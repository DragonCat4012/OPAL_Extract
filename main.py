import os
import re
import shutil
import sys
import zipfile

from lib.Logger import Logger
from lib.team import Team
from lib.xslx_formatter import format_xlsx

imma_nr_map = {}  # read entrys from readme.txt
groups = []  #  all sub-groups

def read_readMe(txt_file_path, imma_nr) -> bool:
    """Extract group member info from readme file"""
    readme_present = False

    if "readme" in f"{txt_file_path}".replace('_', '').strip().lower():# == "readme.txt": # sometimes files named like "readme (1).txt"???
        readme_present = True

        with open(txt_file_path, "r", encoding="utf-8") as file:
            content = file.read()
            seven_digit_numbers = re.findall(r"\b\d{7}\b", content)
            if imma_nr in seven_digit_numbers:
                seven_digit_numbers.remove(imma_nr)
            Logger.info(f"Group members: {seven_digit_numbers}", 1)
            add_team(seven_digit_numbers, imma_nr)
    return readme_present


def extract_zip(zip_file_path, extraction_folder, print_info=False):
    """Extract zip files"""
    os.makedirs(extraction_folder, exist_ok=True)

    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(extraction_folder)

    # print(f"> Successfully extracted to {extraction_folder}")

    rating_file_name = None
    for item in os.listdir(extraction_folder):
        if item.endswith(".xlsx"):
            # Logger.info(item, 1)
            rating_file_name = os.path.join(extraction_folder, item)
    move_extracted_content(extraction_folder, print_info)

    return rating_file_name


def add_team(members: list[str], imma_nr: str):
    """Add immanrs to teams or create new one"""
    for team in groups:
        if team.is_member(imma_nr):
            return

    members.append(imma_nr)
    groups.append(Team(members))


def move_extracted_content(parent_folder, print_info):
    """Move dropboxes and extratc content from them"""
    extracted_dirs = [
        d
        for d in os.listdir(parent_folder)
        if os.path.isdir(os.path.join(parent_folder, d))
    ]

    if not extracted_dirs:
        # print("No folders found to extract.")
        return

    nested_folder = os.path.join(parent_folder, extracted_dirs[0])

    if print_info:
        print("---------------------Dropboxes-------------------------")

    for item in os.listdir(nested_folder):
        if print_info:
            Logger.info_colored(f"{item}")
        imma_nr = f"{item}".split("_")[-1]

        item_path = os.path.join(nested_folder, item)

        # Assignment entzip
        if os.path.isdir(item_path):
            all_files = os.listdir(item_path)
            for file in all_files:
                file_path = os.path.join(item_path, file)

                if file_path.endswith(".zip"):
                    Logger.info(f"Assignment: {file}", 1)
                    extract_zip(file_path, item_path)
                    os.remove(file_path)

                    submission_files = os.listdir(item_path)
                    folder_names = [
                        entry
                        for entry in submission_files
                        if os.path.isdir(os.path.join(item_path, entry))
                    ]
                    submission_files_level2 = []

                    # Extract subfolder if not directly zipped
                    for folder in folder_names:

                        if folder == file.split(".")[0]:
                            #print("\t  Ͱ " + folder)
                            dir_path = os.path.join(item_path, folder)

                            for x in os.listdir(dir_path):
                              #  print("\t  Ͱ " + x)
                                shutil.move(os.path.join(dir_path, x), item_path)
                                submission_files.append(x)
                            os.rmdir(os.path.join(item_path, folder))
                        else:
                            if folder == "__MACOSX": # ignore mac files
                                continue

                            submission_files.remove(folder) # dont print again
                            print("\t  ╙ " + folder)
                            dir_path = os.path.join(item_path, folder)

                            for x in os.listdir(dir_path):
                                print("\t\t  Ͱ " + x)
                                #shutil.move(os.path.join(dir_path, x), item_path)
                                #submission_files.append(x)
                                submission_files_level2.append(x)

                        #    os.rmdir(os.path.join(item_path, folder))

                    # parse group member info
                    readme_present = False
                    if len(submission_files) > 0:
                       print("\t  Ͱ " + "\n\t  Ͱ ".join(submission_files))

                    for sub_file in submission_files + submission_files_level2:
                        txt_file_path = os.path.join(item_path, sub_file)
                        result = read_readMe(txt_file_path, imma_nr)
                        readme_present =  result if result else readme_present # only update to true

                    if not readme_present:
                        Logger.error("No readme found or readme in subfolder:c", 2)
                        add_team([], imma_nr)

        # Nested Zips
        if item.endswith(".zip"):
            extract_zip(item_path, nested_folder, True)

            for extracted_item in os.listdir(nested_folder):
                # print(f"\t\t{extracted_item}")

                extracted_item_path = os.path.join(nested_folder, extracted_item)

                if not "dropboxes" in extracted_item_path:
                    continue
                # print("> Dropboxes:" + extracted_item_path)

                if os.path.isdir(extracted_item_path):
                    for file in os.listdir(extracted_item_path):
                        shutil.move(
                            os.path.join(extracted_item_path, file), parent_folder
                        )
                    os.rmdir(extracted_item_path)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        Logger.error("Usage: python3 main.py <zip_file_path> <groupnumber>")
        sys.exit(1)

    output_zip_file_path = sys.argv[1]
    groupnumber = sys.argv[2]
    FOLDERNAME = f"GruMCI G{groupnumber}"

    # remove dir if already exists
    if os.path.exists(FOLDERNAME):
        shutil.rmtree(FOLDERNAME)
        Logger.info(f"Previous Directory [{FOLDERNAME}] has been deleted successfully.")

    file_path = extract_zip(output_zip_file_path, FOLDERNAME)
    format_xlsx(file_path, groupnumber, groups)

import zipfile
import os
import sys
import shutil

import re

from lib.team import Team
from lib.xslx_formatter import format_xlsx

imma_nr_map = {} # read entrys from readme.txt
groups = [] #  all sub-groups

def extract_zip(zip_file_path, extraction_folder, print_info = False):
    """Extract zip files"""
    os.makedirs(extraction_folder, exist_ok=True)

    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(extraction_folder)

    #print(f"> Successfully extracted to {extraction_folder}")

    rating_file_name = None
    for item in os.listdir(extraction_folder):
        if item.endswith(".xlsx"):
            print(f"\t{item}")
            rating_file_name = os.path.join(extraction_folder, item)
    move_extracted_content(extraction_folder, print_info)

    #print(f"> Returning : {rating_file_name}")
    return rating_file_name


def add_team(members, imma_nr):
    """Add immanrs to teams or create new one"""
    for t in groups:
        if t.is_member(imma_nr):
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
        #print("No folders found to extract.")
        return

    nested_folder = os.path.join(parent_folder, extracted_dirs[0])

    if print_info:
        print("---------------------Dropboxes-------------------------")

    for item in os.listdir(nested_folder):
        if print_info:
            print(f"\t{item}")
        imma_nr = f"{item}".split("_")[-1]

        item_path = os.path.join(nested_folder, item)

        # Assignment entzip
        if os.path.isdir(item_path):
            all_files = os.listdir(item_path)
            for file in all_files:
                file_path = os.path.join(item_path, file)

                if file_path.endswith(".zip"):
                    print(f"\t\tAssignment: {file}")
                    extract_zip(file_path, item_path)

                    all_files2 = os.listdir(item_path)
                    for file2 in all_files2:
                        if f"{file2}".lower() == "readme.txt":
                            print(f"\t\t\t{file2}")
                            txt_file_path = os.path.join(item_path, file2)
                            with open(txt_file_path, 'r', encoding='utf-8') as file:
                                content = file.read()
                                seven_digit_numbers = re.findall(r'\b\d{7}\b', content)
                                if imma_nr in seven_digit_numbers:
                                    seven_digit_numbers.remove(imma_nr)
                                print(f"\t\t\tGroup members: {seven_digit_numbers}")
                                add_team(seven_digit_numbers, imma_nr)

        # Nested Zips
        if item.endswith(".zip"):
            extract_zip(item_path, nested_folder, True)

            for extracted_item in os.listdir(nested_folder):
                print(f"\t\t{extracted_item}")

                extracted_item_path = os.path.join(nested_folder, extracted_item)

                if not "dropboxes" in extracted_item_path:
                    continue
                print("> Dropboxes:" + extracted_item_path)

                if os.path.isdir(extracted_item_path):
                    for file in os.listdir(extracted_item_path):
                        shutil.move(
                            os.path.join(extracted_item_path, file), parent_folder
                        )
                    os.rmdir(extracted_item_path)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <zip_file_path> <groupnumber>")
        sys.exit(1)

    output_zip_file_path = sys.argv[1]
    groupnumber = sys.argv[2]
    FOLDERNAME = f"GruMCI G{groupnumber}"

    # remove dir if already exists
    if os.path.exists(FOLDERNAME):
        shutil.rmtree(FOLDERNAME)
        print("Directory ZIP1 has been deleted successfully.")

    filePath = extract_zip(output_zip_file_path, FOLDERNAME)
    format_xlsx(filePath, groupnumber, groups)

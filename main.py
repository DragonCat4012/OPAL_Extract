import zipfile
import os
import sys
import shutil
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

dropbox_immanrs = []


def extract_zip(zip_file_path, extraction_folder):
    os.makedirs(extraction_folder, exist_ok=True)

    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(extraction_folder)

    print(f"> Successfully extracted to {extraction_folder}")

    rating_file_name = None
    for item in os.listdir(extraction_folder):
        if item.endswith(".xlsx"):
            print(item)
            rating_file_name = os.path.join(extraction_folder, item)
    move_extracted_content(extraction_folder)

    return rating_file_name


def move_extracted_content(parent_folder):
    global dropbox_immanrs
    extracted_dirs = [
        d
        for d in os.listdir(parent_folder)
        if os.path.isdir(os.path.join(parent_folder, d))
    ]

    if not extracted_dirs:
        print("No folders found to extract.")
        return

    nested_folder = os.path.join(parent_folder, extracted_dirs[0])

    print("---------------------Dropboxes-------------------------")

    for item in os.listdir(nested_folder):
        print("\t" + item)
        dropbox_immanrs.append(f"{item}".split("_")[-1])

        item_path = os.path.join(nested_folder, item)

        if item.endswith(".zip"):
            extract_zip(item_path, nested_folder)

            for extracted_item in os.listdir(nested_folder):
                extracted_item_path = os.path.join(nested_folder, extracted_item)

                if not "dropboxes" in extracted_item_path:
                    continue

                print("\t\t" + extracted_item_path)

                if os.path.isdir(extracted_item_path):
                    for file in os.listdir(extracted_item_path):
                        shutil.move(
                            os.path.join(extracted_item_path, file), parent_folder
                        )
                    os.rmdir(extracted_item_path)


def formatXSL(input_file_path, group):
    # print("XSL: " + input_file_path)

    output_file_path = os.path.join(
        input_file_path.split("/")[0], f"Bewertung_{group}.xlsx"
    )

    columns_to_remove = [
        "Anrede",
        "Studiengruppe",
        "Organisationseinheit",
        "Fachsemester",
        "Studiengang",
        "Studienabschluss",
        "Institution",
        "Standort",
        "Startdatum",
        "Dauer",
    ]

    remove_columns_from_xls(input_file_path, output_file_path, columns_to_remove)


def remove_columns_from_xls(input_file_path, output_file_path, columns_to_remove):
    df = pd.read_excel(input_file_path)
    df.drop(columns=columns_to_remove, inplace=True)
    df.to_excel(output_file_path, engine="openpyxl", index=False)

    workbook = load_workbook(output_file_path)
    sheet = workbook.active

    fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")

    header = df.columns.tolist()
    target_index = header.index("Matrikelnummer")

    count = 0

    for row in sheet.iter_rows(min_row=1):
        if f"{row[target_index].value}" in dropbox_immanrs:
            count += 1
            for cell in row:
                cell.fill = fill
    workbook.save(output_file_path)
    print(f"{count} persons marked in excel sheet")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python script.py <zip_file_path> <groupnumber>")
        sys.exit(1)

    zip_file_path = sys.argv[1]
    groupnumber = sys.argv[2]
    foldername = f"GruMCI G{groupnumber}"

    # rmeove dir if already exists
    if os.path.exists(foldername):
        shutil.rmtree(foldername)
        print("Directory ZIP1 has been deleted successfully.")

    filePath = extract_zip(zip_file_path, foldername)
    formatXSL(filePath, groupnumber)

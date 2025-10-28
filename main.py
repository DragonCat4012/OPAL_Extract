import zipfile
import os
import sys
import shutil
import pandas as pd


def extract_zip(zip_file_path, extraction_folder):
    # Create the extraction folder if it does not exist
    os.makedirs(extraction_folder, exist_ok=True)

    # Extract ZIP file
    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(extraction_folder)

    print(f"Successfully extracted to {extraction_folder}")

    retunrstr = None
    for item in os.listdir(extraction_folder):
        if item.endswith(".xlsx"):
            print(item)
            retunrstr = os.path.join(extraction_folder, item)
    move_extracted_content(extraction_folder)

    return retunrstr


def move_extracted_content(parent_folder):
    # Get the list of directories in the parent folder
    extracted_dirs = [
        d
        for d in os.listdir(parent_folder)
        if os.path.isdir(os.path.join(parent_folder, d))
    ]

    if not extracted_dirs:
        print("No folders found to extract.")
        return

    # Assuming there's only one folder in the parent directory
    nested_folder = os.path.join(parent_folder, extracted_dirs[0])

    print("---------------------")

    # Extract any ZIP files in the nested folder
    for item in os.listdir(nested_folder):
        print(item)
        item_path = os.path.join(nested_folder, item)

        if item.endswith(".zip"):
            extract_zip(item_path, nested_folder)
            # Move extracted contents up one level

            for extracted_item in os.listdir(nested_folder):
                extracted_item_path = os.path.join(nested_folder, extracted_item)

                if not "dropboxes" in extracted_item_path:
                    continue

                print("\t" + extracted_item_path)
                if os.path.isdir(extracted_item_path):
                    for file in os.listdir(extracted_item_path):
                        shutil.move(
                            os.path.join(extracted_item_path, file), parent_folder
                        )
                    os.rmdir(extracted_item_path)


def formatXSL(input_file_path):
    print("XSL: " + input_file_path)

    output_file_path = os.path.join(input_file_path.split("/")[0], "Bewertung.xlsx")

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
    df.to_excel(output_file_path, index=False)


def extract_zipOLD(zip_file_path, extraction_folder):
    os.makedirs(extraction_folder, exist_ok=True)

    # Extract ZIP file
    with zipfile.ZipFile(zip_file_path, "r") as zip_ref:
        zip_ref.extractall(extraction_folder)

    print(f"Successfully extracted to {extraction_folder}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python script.py <zip_file_path> <extraction_folder>")
        sys.exit(1)

    # Get arguments
    zip_file_path = sys.argv[1]
    if os.path.exists("ZIP1"):
        # Remove the directory and all its contents
        shutil.rmtree("ZIP1")
        print(f"Directory {'ZIP1'} has been deleted successfully.")

    filePath = extract_zip(zip_file_path, "ZIP1")
    formatXSL(filePath)

import os
from book_of_d_to_google_sheets.google_sheets_handler import GoogleSheetsHandler
from document_reader import DocumentReader


def main():
    document_reader = DocumentReader("book-of-d.docx")
    hazards = document_reader.get_hazards()
    print(hazards)

    google_sheets_handler = GoogleSheetsHandler(
        os.path.expanduser("~/safe-elephant/safe-elephant-412612-0dedd3700ba7.json"), "Book of D - Parsed Data MAIN"
    )

    headers = ["CATEGORY", "HAZARD", "TO WHOM", "CONTROLS"]
    google_sheets_handler.write_headers(headers)

    # Prepare a list to hold all rows before a batch update
    all_rows = []

    for category, subcategories in hazards.items():
        for hazard, details in subcategories.items():
            to_whoms = details.get("To Whom", [])
            controls = details.get("Controls", [])
            
            for to_whom in to_whoms:
                for control in controls:
                    row = [category, hazard, to_whom, control]
                    all_rows.append(row)

            if not to_whoms and not controls:
                row = [category, hazard, '', '']
                all_rows.append(row)

    # Append all rows in a batch
    google_sheets_handler.worksheet.append_rows(all_rows)

if __name__ == "__main__":
    main()

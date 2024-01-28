from document_reader import DocumentReader


def main():
    document_reader = DocumentReader("book-of-d.docx")
    hazards = document_reader.get_hazards()
    print(hazards)

if __name__ == "__main__":
    main()

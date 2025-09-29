import os
from docx2pdf import convert


def convert_word_to_pdf_in_folder(input_folder, output_folder=None):
    """
    Converts all .docx files in a specified input folder to PDF.
    Optionally saves the PDFs to a separate output folder.
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for filename in os.listdir(input_folder):
        if filename.endswith(".docx"):
            docx_path = os.path.join(input_folder, filename)

            if output_folder:
                pdf_filename = filename.replace(".docx", ".pdf")
                pdf_path = os.path.join(output_folder, pdf_filename)
            else:
                # If no output folder specified, save PDF in the same folder
                pdf_path = docx_path.replace(".docx", ".pdf")

            try:
                convert(docx_path, pdf_path)
                print(f"Converted '{filename}' to PDF.")
            except Exception as e:
                print(f"Error converting '{filename}': {e}")


# --- Usage Example ---
if __name__ == "__main__":
    input_directory = "C:\\Users\\MehdiMajid\\Documents\\WordFiles"  # Replace with your input folder path
    output_directory = "C:\\Users\\MehdiMajid\\Documents\\PDFOutput"  # Optional: Replace with your desired output folder path

    # Convert and save PDFs to a separate output folder
    convert_word_to_pdf_in_folder(input_directory, output_directory)

    # Or, convert and save PDFs in the same input folder (comment out the above line if using this)
    # convert_word_to_pdf_in_folder(input_directory)

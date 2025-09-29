import os, subprocess
from pathlib import Path
from docx2pdf import convert

# print(os.path.join("C:\\tmp", "file.docx"))
# print(Path("C:/tmp") / "file.docx")

def convert_pdf_to_docx(input_folder, output_folder=None):
    """
    Converts all .docx files in a specified folder to PDF.
    Optionally saves the PDFs to a separate folder.
    Skips non-.docx files and converts all .docx files to PDF only.
    :param input_folder:
    :param output_folder:
    :return:
    """
    # if output_folder and not os.path.exists(output_folder):
    #     os.makedirs(output_folder)
    #     print(f"Created folder for:  {output_folder}")
    n = 0
    for root, dirs, files in os.walk(input_folder):
        print(f"root: {root}, dirs: {dirs}, files: {files}, len(files)={len(files)}, ")
        print(f"{root} {len(files)}")
        if n >= 0:
            try:
                # result = subprocess.run(['tree',"/A", f"{root}"], capture_output=True, text=True, check=True, shell=True)
                result = subprocess.run(['tree',"/f", f"{root}"], capture_output=True, text=True, check=True, shell=True)
                print(result.stdout)
            except FileNotFoundError:
                print("No such file or directory")
        n+=1
        # print((subprocess.run(['tree "{root}"'], capture_output=True, text=True, check=True, shell=True)).stdout)


# convert_pdf_to_docx(input_folder="C:\\Users\\MehdiMajid\\Documents")
convert_pdf_to_docx(input_folder=Path("C:/Users/MehdiMajid/Documents"))
# rooti = Path("C:/Users/MehdiMajid/Documents/")
# print(rooti)
# print(rooti.exists())
# print(rooti.is_dir())
# print(type(rooti))
# print((subprocess.run(["tree", r"C:/Users/MehdiMajid/Documents/"], capture_output=True, text=True, check=True, shell=True)).stdout)
# print((subprocess.run(["tree", f"{rooti}"], capture_output=True, text=True, check=True, shell=True)).stdout)

def __main__():
    # convert_pdf_to_docx("C:\\Users\\MehdiMajid\\Document")
    convert_pdf_to_docx(Path("C:/Users/MehdiMajid/Documents"))
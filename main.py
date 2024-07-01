from spire.presentation import Presentation, FileFormat
import os
import shutil  # Import shutil for file operations

def convert_ppt_to_pdf(input_file: str, output_file: str) -> None:
    """
    Convert a PowerPoint file to PDF using Spire.Presentation.

    Parameters:
    - input_file (str): The path to the input PowerPoint file.
    - output_file (str): The path where the output PDF file will be saved.

    Returns:
    - None
    """
    try:
        presentation = Presentation()
        presentation.LoadFromFile(input_file)
        presentation.SaveToFile(output_file, FileFormat.PDF)
        presentation.Dispose()
        print(f'Successfully converted {input_file} to {output_file}')
    except Exception as e:
        print(f"Error converting presentation: {e}")

def organize_and_convert_files(folder: str) -> None:
    """
    Organize PPT/PPTX files into an 'inputs' directory and convert them to PDF,
    saving the PDFs in an 'output' directory. Skips conversion if PDF already exists.

    Parameters:
    - folder (str): The path to the folder containing PowerPoint files to be organized and converted.

    Returns:
    - None
    """
    input_folder = os.path.join(folder, "inputs")
    output_folder = os.path.join(folder, "output")

    # Create 'inputs' and 'output' directories if they don't exist
    os.makedirs(input_folder, exist_ok=True)
    os.makedirs(output_folder, exist_ok=True)

    # Move PPT/PPTX files to 'inputs' directory
    for file in os.listdir(folder):
        if file.endswith((".ppt", ".pptx")):
            src_path = os.path.join(folder, file)
            dest_path = os.path.join(input_folder, file)
            if not os.path.exists(dest_path):  # Move only if it doesn't already exist in 'inputs'
                shutil.move(src_path, dest_path)

    # Convert files in 'inputs' to PDF in 'output', skipping if PDF already exists
    for pptfile in os.listdir(input_folder):
        input_file = os.path.join(input_folder, pptfile)
        output_file = os.path.join(output_folder, os.path.splitext(pptfile)[0] + ".pdf")
        if not os.path.exists(output_file):  # Convert only if the PDF doesn't already exist
            convert_ppt_to_pdf(input_file, output_file)
        else:
            print(f"Skipping conversion for {pptfile}; PDF already exists.")

if __name__ == "__main__":
    cwd = os.getcwd()
    organize_and_convert_files(cwd)
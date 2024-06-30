from spire.presentation import Presentation, FileFormat
import os

def convert_ppt_to_pdf(input_file, output_file):
    """Convert a PowerPoint file to PDF using Spire.Presentation."""
    try:
        presentation = Presentation()
        presentation.LoadFromFile(input_file)
        presentation.SaveToFile(output_file, FileFormat.PDF)
        presentation.Dispose()
        print(f'Successfully converted {input_file} to {output_file}')
    except Exception as e:
        print(f"Error converting presentation: {e}")

def convert_files_in_folder(folder):
    """Convert all PPT/PPTX files in the given folder to PDF."""
    files = os.listdir(folder)
    pptfiles = [f for f in files if f.endswith((".ppt", ".pptx"))]
    for pptfile in pptfiles:
        input_file = os.path.join(folder, pptfile)
        output_file = os.path.splitext(input_file)[0] + ".pdf"
        convert_ppt_to_pdf(input_file, output_file)

if __name__ == "__main__":
    cwd = os.getcwd()
    convert_files_in_folder(cwd)

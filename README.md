# PowerPoint to PDF Conversion Tool

This Python script automates the conversion of PowerPoint (.ppt and .pptx) files to PDF format using Spire.Presentation. It scans a specified folder for PowerPoint files and converts each one to a PDF file in the same location.

## Prerequisites

- Python 3.9
- Spire.Presentation for Python

## Installation

First, ensure that Python 3.9 is installed on your system. Then, install Spire.Presentation via pip:

```sh

pip install Spire.Presentation
```

Alternatively, you can create a conda environment using the provided [`environment.yaml`](vscode-file://vscode-app/c:/Program%20Files/Microsoft%20VS%20Code/resources/app/out/vs/code/electron-sandbox/workbench/workbench.html "environment.yaml") file:

```
conda env create -f environment.yaml
```

## Usage

To use the script, simply run `main.py` from your terminal or command prompt:

```
python main.py
```

The script will automatically find and convert all PowerPoint files in the current working directory to PDF format.

## How It Works

The script uses the `Spire.Presentation` library to open each PowerPoint file and save it as a PDF. It looks for files with the `.ppt` or `.pptx` extension in the specified folder and processes each file found.

## Contributing

Contributions to this project are welcome. Please feel free to fork the repository, make your changes, and submit a pull request.

## License

This project is open-source and available under the MIT License.

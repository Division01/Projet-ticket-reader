# Webcam Text Reader

### Author
Vincent Fischer

## Overview

`webcam_text_reader.py` is a Python script that captures images from a webcam feed or processes pre-existing image files to extract text using Optical Character Recognition (OCR). The extracted text is saved to both a `.txt` file and an Excel (`.xlsx`) file.
This code was made to work on Colruyt receipts. It was custom made for the needs of a project, and as such may not work for other purposes. Feel free to use it as a starting point for your own projects, maybe to extract text from receipts of other stores. It's only important to note that the code is made to work with the format of the receipts, and may need to be changed for other formats.
I did clean the code a bit, but I didn't have any tickets to try it, meaning that it may not work for now. I will try to fix the bugs whenever I can (which may not be soon).


## Versions

This project now has two versions:

1. **Tesseract Version**: Uses Tesseract OCR (requires separate installation)
2. **EasyOCR Version**: Uses EasyOCR (easier to set up, especially on Windows)

## Requirements

1. Python 3.x
2. Libraries:
   - opencv-python
   - openpyxl

### For Tesseract Version:
   - pytesseract
   - Tesseract OCR installed and configured (can be challenging on Windows)

### For EasyOCR Version:
   - easyocr

## Installation

1. Clone this repository.
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

### For Tesseract Version:
3. Install Tesseract OCR and ensure it's properly configured.

### For EasyOCR Version:
3. No additional installation required.

## Usage

### Running the Script

To run the script, use the following command:

```bash
python webcam_text_reader.py
```

### Options

- **Webcam Mode**: By default, the script will attempt to access your webcam to capture images for text extraction.
- **Image File Mode**: To process pre-existing image files, provide the file path as an argument:

  ```bash
  python webcam_text_reader.py --image path/to/image.jpg
  ```

### Output

- **Text File**: The extracted text will be saved in a `.txt` file in the same directory as the script.
- **Excel File**: The extracted text will also be saved in an Excel file (`.xlsx`) for easy viewing and editing.

## Choosing a Version

- If you're on Windows or prefer an easier setup, use the EasyOCR version.
- If you need specific features of Tesseract or are comfortable with its installation, use the Tesseract version.

## Troubleshooting

- **Tesseract Version**: Ensure that Tesseract OCR is installed and the path to the Tesseract executable is correctly set in your system's PATH or specified in the script.
- **EasyOCR Version**: Make sure you have sufficient disk space and a stable internet connection for the initial download of language models.
- **Webcam Access Issues**: Verify that your webcam is properly connected and not being used by another application.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

## Contact

For any questions or issues, please contact Vincent Fischer on LinkedIn or GitHub.
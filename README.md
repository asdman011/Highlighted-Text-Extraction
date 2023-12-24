# Highlighted Text Extraction App
#### Video Demo:  <https://www.youtube.com/watch?v=FOceFzhiMVs>
This Flask application allows you to upload a DOCX file, extract highlighted text with corresponding colors, and download the results in a CSV file.

## Features

- Upload a DOCX file.
- Extract highlighted text with color information.
- View the highlighted text in a tabular format.
- Download the extracted data as a CSV file.

## Installation

1. Clone the repository:

    git clone https://github.com/asdman011/Highlighted-Text-Extraction.git

2. Install dependencies:

    pip install -r requirements.txt


3. Run the application:


    python app.py
or
    flask run


4. Open your browser and visit [http://localhost:5000](http://localhost:5000) to use the application.

## Usage

1. Upload a DOCX file using the file input.
2. Click the "Extract Text" button to process the file.
3. View the extracted highlighted text in a table.
4. Download the highlights as a CSV file.

## Folder Structure

- `uploads/`: Temporary storage for uploaded DOCX files.
- `outputs/`: Destination folder for generated CSV files.

## Dependencies

- Flask
- python-docx
- csv

## Author

Ammar Ashour


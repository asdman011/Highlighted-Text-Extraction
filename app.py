from flask import Flask, render_template, request, send_file, session
import docx
import csv
import os

app = Flask(__name__)
app.secret_key = "1111"  # Add a secret key for session management

def get_color_name(highlight_color):
    # Check if the highlight_color is an instance of the enum
    if isinstance(highlight_color, docx.enum.text.WD_COLOR_INDEX):
        # Convert the enum value to a string
        return highlight_color.name
    else:
        return None

def extract_highlighted_text(docx_file):
    
    # Open the input file as a docx document object
    doc = docx.Document(docx_file)

    # Create a dictionary to store highlighted text by color
    highlighted_data = {}

    # Loop through all the paragraphs in the document
    for para in doc.paragraphs:
        # Loop through all the runs in each paragraph
        for run in para.runs:
            # Check if the run has a highlight color
            highlight_color = run.font.highlight_color
            color_name = get_color_name(highlight_color)

            if color_name:
                # Add highlighted text to the dictionary under the color key
                highlighted_data.setdefault(color_name, []).append(run.text)

    return highlighted_data

@app.route('/', methods=['GET', 'POST'])
def index():
    # Clear session data on page reload
    session.clear()

    if request.method == 'POST':
        docx_file = request.files['docxFile']
        if docx_file and docx_file.filename.endswith('.docx'):
            # Save the uploaded file to a temporary location
            temp_path = os.path.join('uploads', docx_file.filename)
            docx_file.save(temp_path)

            # Process the uploaded file
            highlighted_data = extract_highlighted_text(temp_path)

            # Remove the temporary uploaded file
            os.remove(temp_path)

            # Store data in the session
            session['highlighted_data'] = highlighted_data

            return render_template('index.html', highlighted_data=highlighted_data)

    return render_template('index.html')

@app.route('/download')
def download_file():
    # Open the output file as a csv writer object
    output_file = os.path.join('outputs', 'highlighted_text.csv')
    with open(output_file, "w", newline="", encoding="utf-8") as csv_file:
        fieldnames = ['Color', 'Highlighted Text']
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL, lineterminator='\n')

        # Write header
        writer.writeheader()

        # Write rows in chunks
        chunk_size = 10000  # Set an appropriate chunk size
        highlighted_data = session.get('highlighted_data', {})

        for color, text_list in highlighted_data.items():
            for i in range(0, len(text_list), chunk_size):
                chunk = text_list[i:i + chunk_size]
                for text in chunk:
                    writer.writerow({'Color': color, 'Highlighted Text': text})

    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    os.makedirs('outputs', exist_ok=True)
    app.run(debug=True)

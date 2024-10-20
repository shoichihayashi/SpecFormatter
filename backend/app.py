from flask import Flask, request, jsonify
from flask_cors import CORS
from formatter import Formatter

app = Flask(__name__)

# enable CORS for ross-origin requests
CORS(app)

@app.route('/process', methods=['POST'])
def process_document():
    if 'source_file' not in request.files or 'target_file' not in request.files:
            return jsonify({'error': 'Files not provided'}), 400
    
    # get uploaded files
    source_file = request.files['source_file']
    target_file = request.files['target_file']

    # call Formatter from class script
    formatter = Formatter(source_file, target_file)

    for source_paragraph, target_paragraph in zip(formatter.source_doc.paragraphs, formatter.target_doc.paragraphs):
        formatter.apply_paragraph_formatting(source_paragraph, target_paragraph)

    formatter.apply_header()
    formatter.apply_footer()

    # save file
    formatter.save_target('modified_target.docx')

    # return success message
    return jsonify({'message': 'Files processed successfully', 'download_link': '/download/modified_target.docx'})

if __name__ == '__main__':
     app.run(debug=True)
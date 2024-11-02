from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from formatter import Formatter
import os
import zipfile
import io
import tempfile
import shutil

app = Flask(__name__)

# enable CORS for cross-origin requests
CORS(app)

@app.route('/process', methods=['POST'])
def process_document():
     # make temp directory
     temp_dir = tempfile.mkdtemp()
     try:        
          # check if source file exist
          if 'source_file' not in request.files:
                print("Source file not provided")
                return jsonify({'error': 'Source file not provided'}), 400
          
          # get uploaded files
          source_file = request.files['source_file']
          target_files = [file for key, file in request.files.items() if key.startswith('target_file')]

          # check if target files exist
          if not target_files:
                print("Target file(s) not provided")
                return jsonify({'error': 'Target file(s) not provided'}), 400

          # target file list
          target_file_paths = []

          # format each target file
          for target_file in target_files:
               print('target_file: ', target_file)
               # call Formatter from class script
               formatter = Formatter(source_file, target_file)

               for source_paragraph, target_paragraph in zip(formatter.source_doc.paragraphs, formatter.target_doc.paragraphs):
                    formatter.apply_paragraph_formatting(source_paragraph, target_paragraph)

               formatter.apply_header()
               formatter.apply_footer()

               # save file
               target_file_path = os.path.join(temp_dir, f'{target_file.filename}')
               formatter.save_target(target_file_path)
               print(f"File saved as {target_file_path}")

               # add to target file name list
               target_file_paths.append(target_file_path)

          zip_archive = io.BytesIO()
          with zipfile.ZipFile(zip_archive, "w") as zip_file:
               for file_path in target_file_paths:
                    zip_file.write(file_path, arcname=os.path.basename(file_path))

          zip_archive.seek(0)

          # return success message
          return send_file(
               zip_archive,
               mimetype="application/zip",
               as_attachment=True,
               download_name="formatted_files.zip"
          )
     except Exception as e:
          print(f"Error: {e}")
          return jsonify({'error': 'Internal server error'}), 500

     # delete the temporary directory
     finally:
          shutil.rmtree(temp_dir)

# route to serve the file for download
@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
     # ensure file exists in backend folder
     file_path = os.path.join(os.getcwd(), filename)
     if os.path.exists(file_path):
          return send_file(file_path, as_attachment=True)
     else:
          return jsonify({'error': 'File not found'}), 404


if __name__ == '__main__':
     app.run(debug=True)
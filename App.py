import os
import pandas as pd
from flask import Flask, request, send_file, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Supported image formats
image_formats = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.j2k', '.jp2')

def get_image_data(directory):
    image_data = []

    # Walk through the directory and its subdirectories
    for root, dirs, files in os.walk(directory):
        # Split the root path into parts (folders)
        folder_path = os.path.relpath(root, directory).split(os.sep)
        
        for file_name in files:
            if file_name.lower().endswith(image_formats):
                # Create a row with folder path split into columns, and image file name
                row = folder_path + [file_name]
                image_data.append(row)
    
    return image_data

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/process_directory', methods=['POST'])
def process_directory():
    # Get the directory path from the form
    directory = request.form['directory']

    if not os.path.isdir(directory):
        return "Invalid directory. Please check the path.", 400

    # Get image data (folder hierarchy + image files)
    image_data = get_image_data(directory)

    # Sort the image_data in ascending order
    image_data.sort(key=lambda x: [str(part).lower() for part in x])  # Sort based on folder/file names, case-insensitive

    # Create DataFrame for image data and add columns dynamically based on folder depth
    max_depth = max(len(row) for row in image_data)  # Find the maximum folder depth
    columns = [f'Column{i+1}' for i in range(max_depth)]  # Create column names like Column1, Column2...

    df_images = pd.DataFrame(image_data, columns=columns)

    # Create a dictionary to count images in each folder
    folder_counts = {}

    for row in image_data:
        folder_path = tuple(row[:-1])  # Get the folder path (all but the last element)
        folder_counts[folder_path] = folder_counts.get(folder_path, 0) + 1

    # Prepare data for the counts DataFrame
    count_data = [(os.path.sep.join(folder), count) for folder, count in folder_counts.items()]

    # Create a DataFrame for folder counts
    df_counts = pd.DataFrame(count_data, columns=['Folder Name', 'Image Count'])

    # Specify output Excel file path
    output_excel = os.path.join(directory, "image_file_structure.xlsx")

    # Write both DataFrames to separate sheets in the Excel file
    with pd.ExcelWriter(output_excel) as writer:
        df_images.to_excel(writer, sheet_name='Image Data', index=False)  # Sheet 1: Detailed Image Data
        df_counts.to_excel(writer, sheet_name='Folder Summary', index=False)  # Sheet 2: Folder Summary with Image Counts

    # Provide the generated Excel file for download
    return send_file(output_excel, as_attachment=True)




if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=False)


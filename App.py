import os
import pandas as pd
from flask import Flask, request, send_file, render_template
import io

app = Flask(__name__)

# Supported image formats
image_formats = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.j2k', '.jp2')

# Maximum rows per Excel sheet (Excel's limit is 1,048,576 rows)
MAX_ROWS_PER_SHEET = 1048576

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

    # Prepare data for the counts DataFrame with rearranged columns
    count_data = [
        (os.path.join(directory, *folder), os.path.sep.join(folder), count)  # Rearranged: Full Path, Folder Name, Image Count
        for folder, count in folder_counts.items()
    ]

    # Create a DataFrame for folder counts with updated column order
    df_counts = pd.DataFrame(count_data, columns=['Full Path', 'Folder Name', 'Image Count'])  # Column order changed

    # Save the Excel file in memory using BytesIO
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Handle large data by writing in chunks to multiple sheets if necessary
        total_rows = df_images.shape[0]
        sheet_num = 1
        for start_row in range(0, total_rows, MAX_ROWS_PER_SHEET):
            # Define the end row for the current sheet
            end_row = min(start_row + MAX_ROWS_PER_SHEET, total_rows)
            # Extract the chunk of data
            df_chunk = df_images.iloc[start_row:end_row]
            # Write chunk to a new sheet (e.g., 'Image Data 1', 'Image Data 2', ...)
            sheet_name = f'Image Data {sheet_num}'
            df_chunk.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_num += 1

        # Write the Folder Summary to a separate sheet
        df_counts.to_excel(writer, sheet_name='Folder Summary', index=False)

    # Rewind the buffer
    output.seek(0)

    # Send the Excel file as a download
    return send_file(output, as_attachment=True, download_name='image_file_structure.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)

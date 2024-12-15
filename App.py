import os
import pandas as pd
from flask import Flask, request, send_file, render_template
import io
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)

# Supported image formats
image_formats = ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.j2k', '.jp2')

# Maximum rows per Excel sheet (Excel's limit is 1,048,576 rows)
MAX_ROWS_PER_SHEET = 1048576

def get_image_data(directory):
    image_data = []
    for root, dirs, files in os.walk(directory):
        folder_path = os.path.relpath(root, directory).split(os.sep)
        for file_name in files:
            if file_name.lower().endswith(image_formats):
                row = folder_path + [file_name]
                image_data.append(row)
    return image_data

def get_folder_size_parallel(folder_path):
    """Calculate folder size using multithreading for faster performance."""
    total_size = 0

    def calculate_size(path):
        try:
            if os.path.isfile(path):
                return os.path.getsize(path)
            elif os.path.isdir(path):
                return get_folder_size_parallel(path)
        except (PermissionError, FileNotFoundError):
            return 0
        return 0

    with ThreadPoolExecutor() as executor:
        futures = [executor.submit(calculate_size, entry.path) for entry in os.scandir(folder_path)]
        total_size += sum(future.result() for future in futures)

    return total_size


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/process_directory', methods=['POST'])
def process_directory():
    directory = request.form['directory']
    action = request.form['action']  # 'full', 'summary', or 'summary_with_size'

    if not os.path.isdir(directory):
        return "Invalid directory. Please check the path.", 400

    image_data = get_image_data(directory)
    image_data.sort(key=lambda x: [str(part).lower() for part in x])

    if action == 'full':
        max_depth = max(len(row) for row in image_data)
        columns = [f'Column{i+1}' for i in range(max_depth)]
        df_images = pd.DataFrame(image_data, columns=columns)

    folder_counts = {}
    folder_sizes = {}

    for row in image_data:
        folder_path = tuple(row[:-1])
        folder_full_path = os.path.join(directory, *folder_path)
        folder_counts[folder_path] = folder_counts.get(folder_path, 0) + 1

        if action == 'summary_with_size' and folder_path not in folder_sizes:
            folder_sizes[folder_path] = get_folder_size_parallel(folder_full_path)

    count_data = []
    for folder, count in folder_counts.items():
        folder_summary = [
            os.path.join(directory, *folder),
            os.path.sep.join(folder),
            count
        ]
        if action == 'summary_with_size':
            folder_summary.append(round(folder_sizes.get(folder, 0) / (1024 ** 3), 2))  # Size in GB
        count_data.append(folder_summary)

    columns = ['Full Path', 'Folder Name', 'Image Count']
    if action == 'summary_with_size':
        columns.append('Folder Size (GB)')

    df_counts = pd.DataFrame(count_data, columns=columns)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if action == 'full':
            total_rows = df_images.shape[0]
            sheet_num = 1
            for start_row in range(0, total_rows, MAX_ROWS_PER_SHEET):
                end_row = min(start_row + MAX_ROWS_PER_SHEET, total_rows)
                df_chunk = df_images.iloc[start_row:end_row]
                df_chunk.to_excel(writer, sheet_name=f'Image Data {sheet_num}', index=False)
                sheet_num += 1

        df_counts.to_excel(writer, sheet_name='Folder Summary', index=False)

    output.seek(0)
    return send_file(output, as_attachment=True, download_name='image_file_structure.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == '__main__':
    app.run(host='127.0.0.1', port=5000, debug=True)

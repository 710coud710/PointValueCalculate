from flask import Flask, request, render_template, redirect, url_for, send_file
import pandas as pd
import os
from fuzzywuzzy import fuzz

app = Flask(__name__)

# Define folders for uploads and processed data
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Create necessary directories if they don't exist
for folder in [UPLOAD_FOLDER, PROCESSED_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

def is_similar(keyword1, keyword2, threshold=90):
    return fuzz.ratio(keyword1.lower(), keyword2.lower()) >= threshold

def are_values_equal(row1, row2):
    return row1[1:].equals(row2[1:])

def filter_duplicate_keywords(data):
    unique_keywords = []
    added_keywords = set()  

    for idx1, row1 in data.iterrows():
        keyword1 = row1['Keyword']
        if keyword1 in added_keywords:
            continue  
        
        similar_found = False

        for unique_row in unique_keywords:
            keyword_unique = unique_row['Keyword']
            if is_similar(keyword1, keyword_unique):          
                if are_values_equal(row1, unique_row):
                    similar_found = True
                    added_keywords.add(keyword1) 
                    break
        if not similar_found:
            unique_keywords.append(row1)
            added_keywords.add(keyword1)

    return pd.DataFrame(unique_keywords)

def calculate_percentile(data, column_name):
    sorted_data = sorted(data[column_name])
    n = len(sorted_data)
    percentiles = []
    for value in data[column_name]:
        r = sorted_data.index(value) + 1  
        k = ((r-0.5) / n) * 100
        percentiles.append(k)
    return percentiles

# Calculate Point Value
def calculate_point_value(data):
    search_value = data['Search Volume Percentile'] * 1
    cpc_value = 100 - (data['CPC Percentile'] * 1)
    competition_value = 100 - (data['Competition Percentile'] * 1)
    data['Point Value'] = (search_value + cpc_value + competition_value) / 3
    return data

def search_keyword(data, keyword, threshold=40):
    results = []
    for idx, row in data.iterrows():
        if is_similar(row['Keyword'], keyword, threshold):
            results.append(row)
    print(f"Kết quả tìm kiếm cho từ khóa '{keyword}': {results}")
    return pd.DataFrame(results)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/calculate', methods=['GET', 'POST'])
def calculate_view():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
   
        if file:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            data = pd.read_excel(filepath)

            filtered_data = filter_duplicate_keywords(data)

            required_columns = ['Search Volume (Global)', 'CPC (Global)', 'Competition (Global)']
            if not all(column in filtered_data.columns for column in required_columns):
                return "Không tìm thấy các cột cần thiết", 400

            # Calculate percentiles
            filtered_data['Search Volume Percentile'] = calculate_percentile(filtered_data, 'Search Volume (Global)')
            filtered_data['CPC Percentile'] = calculate_percentile(filtered_data, 'CPC (Global)')
            filtered_data['Competition Percentile'] = calculate_percentile(filtered_data, 'Competition (Global)')

            # Calculate point values
            final_data = calculate_point_value(filtered_data)

            processed_filepath = os.path.join(app.config['PROCESSED_FOLDER'], 'calculated_data.xlsx')
            final_data.to_excel(processed_filepath, index=False)

            return redirect(url_for('calculate_view'))

    elif request.method == 'GET':
        sort_column = request.args.get('sort_column', 'Point Value')
        order = request.args.get('order', 'desc')
        keyword = request.args.get('search_keyword', '').strip()

        processed_filepath = os.path.join(app.config['PROCESSED_FOLDER'], 'calculated_data.xlsx')
        if not os.path.exists(processed_filepath):
            return redirect(url_for('index'))

        data = pd.read_excel(processed_filepath)
        display_data = data.copy()
        message = ''

        # Search
        if keyword:
            display_data = search_keyword(display_data, keyword, threshold=50)
            if display_data.empty:
                message = f"Không tìm thấy kết quả nào cho từ khóa '{keyword}'."
                return render_template('result.html', tables='', message=message, search_keyword=keyword)

        # Sort
        if sort_column in display_data.columns:
            ascending = True if order == 'asc' else False
            display_data = display_data.sort_values(by=sort_column, ascending=ascending)

        # Prepare data for display
        display_data = display_data[['Keyword', 'Point Value', 'Search Volume (Global)', 'CPC (Global)', 'Competition (Global)', 'Trending %']]
        #display_data = display_data
        table_html = display_data.to_html(classes='data', index=False)

        return render_template('result.html', tables=table_html, message=message, search_keyword=keyword, sort_column=sort_column, order=order)

    return redirect(url_for('index'))

@app.route('/search', methods=['POST'])
def search_view():
    keyword = request.form.get('search_keyword', '').strip()
    if keyword:
        return redirect(url_for('calculate_view', search_keyword=keyword))
    return redirect(url_for('calculate_view'))

@app.route('/download', methods=['GET'])
def download_file():
    keyword_limit = request.args.get('keyword_limit_select', 'all')  # Get the selected limit
    filepath = os.path.join(app.config['PROCESSED_FOLDER'], 'calculated_data.xlsx')

    if os.path.exists(filepath):
        data = pd.read_excel(filepath)
        
        # If user selected a specific limit, limit the number of rows
        if keyword_limit != 'all':
            keyword_limit = int(keyword_limit)
            data = data.head(keyword_limit)
        
        # Save the limited data to a new file
        limited_filepath = os.path.join(app.config['PROCESSED_FOLDER'], f'calculated_data_{keyword_limit}.xlsx')
        data.to_excel(limited_filepath, index=False)

        return send_file(limited_filepath, as_attachment=True, download_name=f'calculated_data_{keyword_limit}.xlsx')
    else:
        return "File không tồn tại", 404


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
  
    
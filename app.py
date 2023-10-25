from flask import Flask, send_from_directory, render_template, request, jsonify
import pandas as pd
import openpyxl

app = Flask(__name__, static_folder='static')

@app.route('/download_excel')
def download_excel():
    return send_from_directory('assets', 'SEMI_data.xlsx', as_attachment=True)

@app.route('/display_excel')
def display_excel():
    # Read the Excel file
    df = pd.read_excel('assets/SEMI_data.xlsx')

    # Convert the DataFrame to HTML
    table_html = df.to_html(classes='excel-table', border=0)

    # Render the HTML page and pass the table_html to it
    return render_template('index.html', table_html=table_html)

@app.route('/update_excel', methods=['POST'])
def update_excel():
    # Get the data from the request
    data = request.json.get('data')
    
    # Load the Excel file
    file_path = 'assets/SEMI_data.xlsx'
    book = openpyxl.load_workbook(file_path)
    sheet = book.active

    # Update cell B2
    sheet['B2'] = data
    
    # Save the updated Excel file
    book.save(file_path)

    # Convert the updated Excel file to a DataFrame and then to HTML
    df = pd.read_excel(file_path)
    updated_table = df.to_html(classes='excel-table', border=0)

    # Return the updated table
    return jsonify(updatedTable=updated_table)

if __name__ == '__main__':
    app.run(debug=True)

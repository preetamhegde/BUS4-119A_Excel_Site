from flask import Flask, send_from_directory, render_template, request, jsonify
import pandas as pd
import openpyxl
from fuzzywuzzy import process


app = Flask(__name__, static_folder='static')


@app.route('/download_excel')
def download_excel():
    return send_from_directory('assets', 'SEMI_data.xlsx', as_attachment=True)

@app.route('')
def display_excel():
    # Read the Excel file
    df = pd.read_excel('assets/SEMI_data.xlsx')

    # Convert the DataFrame to HTML
    table_html = df.to_html(classes='excel-table', border=0)

    # Render the HTML page and pass the table_html to it
    return render_template('index.html', table_html=table_html)



@app.route('/update_excel', methods=['POST'])
def update_excel():
    metrics_to_columns = {
    "Total water consumed": "B",
    "Municipal water usage": "C",
    "Surface water usage": "D",
    "Groundwater usage": "E",
    "Water restored": "F",
    "Water reclaimed/reused": "G",
    "Water discharged": "H",
    "Quality of water discharged": "I",
    "Non-hazardous waste generated": "J",
    "Hazardous waste generated": "K",
    "Waste recycled (onsite or offsite)": "L",
    "Waste sent to landfill (hazardous and non-hazardous)": "M",
    "Waste incinerated (also referred to as 'energy recovery')": "N"
    }
    
    data = request.json.get('data')
    formatted_data = jsonifyData(data)
    
    file_path = 'assets/SEMI_data.xlsx'
    book = openpyxl.load_workbook(file_path)
    sheet = book.active





    company_name = formatted_data.get("Company Name")

    companies = [sheet.cell(row=row, column=1).value for row in range(1, sheet.max_row + 1)]
    
    # Use fuzzy matching to find the closest match to the given company name
    closest_company_name = find_best_match(company_name, companies)

    print(company_name)
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value == closest_company_name:
            company_row = row
            break

    for key, value in formatted_data.items():
        if key in metrics_to_columns:
          col = metrics_to_columns[key]
          cell = f"{col}{company_row}"
          sheet[cell] = value

   

    book.save(file_path)

    # Convert the updated Excel file to a DataFrame and then to HTML
    df = pd.read_excel(file_path)
    updated_table = df.to_html(classes='excel-table', border=0)

    # Return the updated table
    return jsonify(updatedTable=updated_table)

def find_best_match(name, choices):
    best_match = process.extractOne(name, choices)
    return best_match[0]  # returns the best match

def jsonifyData(data):
    lines = data.split("\n")
    result = {}
    company = lines[0].strip()
    
    result["Company Name"] = company  

    for line in lines[2:]:
        if ": " in line: 
            key, value = line.split(": ", 1)
            result[key] = value

    return result


if __name__ == '__main__':
    app.run(debug=True)

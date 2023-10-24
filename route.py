from flask import Flask, send_from_directory, render_template
import pandas as pd

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

if __name__ == '__main__':
    app.run(debug=True)

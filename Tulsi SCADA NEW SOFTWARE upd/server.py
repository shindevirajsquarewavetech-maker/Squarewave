import pyodbc
import pandas as pd
import io
import os
import subprocess
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape, A0
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.graphics.shapes import Drawing
from reportlab.graphics.charts.barcharts import HorizontalBarChart

app = Flask(__name__)
CORS(app)

SQL_CONFIG = {
    'server': 'DESKTOP-42HFTV6\\SQLEXPRESS', 
    'database': 'Viraj', 
    'username': '', 
    'password': ''
}

CHEMICALS = [
    'AC1', 'AC2', 'AC10', 'AC6', 'AC5', 'RT1', 'AC7', 'SAKHAR', 'AC20', 'AC12', 
    'AO1', 'AC9', 'AC4', 'CAT', 'PIG54_WS', 'R7', 'MWAX', 'AC11', 'R421', 'PIG55', 
    'R222', 'CHEM13', 'DAID', 'MIRA', 'CAKES', 'CHEM1068', 'Che_PIG_54ACT', 'CHEM80', 
    'CA_OH_2', 'CHEM50', 'CHEM4000', 'PIG_55_VITON', 'RN', 'GUD', 'AC13', 'RED_OXIDE', 
    'AC14', 'AC15', 'OIL_NO_4', 'AC16', 'BLUE_COLOUR', 'AC8', 'RED_COLOUR', 'YELLOW_COLOUR', 
    'CDPA', 'S_STERATE', 'P_STERATE', 'OIL_NO_15', 'P_JELLY', 'ZBT', 'EB_DUST', 'ATO', 
    'ACC_PIG_54_ACT'
]

def get_db_connection():
    conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SQL_CONFIG['server']};DATABASE={SQL_CONFIG['database']};Trusted_Connection=yes;"
    return pyodbc.connect(conn_str)

def build_query(start_date, end_date):
    cols = "RecordTime, CycleStart, RECIPE_CHE, RECIPE_ACC, " + ", ".join(CHEMICALS) + ", CycleEnd"
    query = f"SELECT {cols} FROM Scada WHERE CAST(RecordTime AS DATE) >= ? AND CAST(RecordTime AS DATE) <= ?"
    params = [start_date, end_date]
    return query, params

def process_report_data(df, report_type='main'):
    if df.empty:
        return df, []
        
    for col in CHEMICALS:
        if col in df.columns:
            # CONVERT FROM GRAMS TO KG
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 1000.0
            
    existing_chems = [c for c in CHEMICALS if c in df.columns]
    
    # Keep chemical column ONLY if it was used at least once
    active_chemicals = [col for col in existing_chems if df[col].sum() > 0]
    
    if report_type == 'monthly':
        df['Month'] = pd.to_datetime(df['CycleStart'], errors='coerce').dt.strftime('%b-%Y')
        
        # Count total batches AND get the target recipe weight
        agg_dict = {'CycleStart': 'count'}
        for chem in active_chemicals:
            agg_dict[chem] = 'max' 
            
        df_grouped = df.groupby(['Month', 'RECIPE_CHE']).agg(agg_dict).reset_index()
        df_grouped.rename(columns={'CycleStart': 'Total Batches', 'RECIPE_CHE': 'Compound Code'}, inplace=True)
        final_cols = ['Month', 'Compound Code', 'Total Batches'] + active_chemicals
        
        df_final = df_grouped[final_cols].copy()
        
        # Calculate Monthly Total: Total Batches * Target Recipe Weight
        for chem in active_chemicals:
            df_final[chem] = round(df_final['Total Batches'] * df_final[chem], 3)
        
        # Add Total Weight Row
        sum_row = {col: '' for col in final_cols}
        sum_row['Month'] = 'TOTAL WEIGHT IN KG'
        for chem in active_chemicals:
            sum_row[chem] = round(df_final[chem].sum(), 3)
            
        df_final = pd.concat([df_final, pd.DataFrame([sum_row])], ignore_index=True)
        
        return df_final, final_cols
        
    else:
        # main dashboard or daily report
        base_start = ['RecordTime', 'CycleStart', 'RECIPE_CHE', 'RECIPE_ACC']
        base_end = ['CycleEnd']
        final_cols = [c for c in base_start if c in df.columns] + active_chemicals + [c for c in base_end if c in df.columns]
        
        df_final = df[final_cols].copy()
        
        # Add Total Weight Row
        sum_row = {col: '' for col in final_cols}
        if 'CycleStart' in final_cols:
            sum_row['CycleStart'] = 'TOTAL WEIGHT IN KG'
        elif 'RecordTime' in final_cols:
            sum_row['RecordTime'] = 'TOTAL WEIGHT IN KG'
            
        for chem in active_chemicals:
            df_final[chem] = df_final[chem].round(3)
            sum_row[chem] = round(df_final[chem].sum(), 3)
            
        df_final = pd.concat([df_final, pd.DataFrame([sum_row])], ignore_index=True)
        
        return df_final, final_cols

@app.route('/')
def home(): return send_file('index.html')

@app.route('/style.css')
def style(): return send_file('style.css')

@app.route('/api/open-downloads', methods=['POST'])
def open_downloads():
    try:
        if os.name == 'nt':
            downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
            subprocess.Popen(f'explorer "{downloads_path}"')
            return jsonify({'status': 'success'})
    except Exception as e:
        pass
    return jsonify({'status': 'ignored'})

@app.route('/api/report', methods=['POST'])
def generate_report():
    try:
        data = request.json
        start_date, end_date = data.get('startDate'), data.get('endDate')
        query, params = build_query(start_date, end_date)
        conn = get_db_connection()
        df = pd.read_sql(query, conn, params=params)
        conn.close()

        if df.empty:
            return jsonify({'status': 'success', 'tableData': [], 'chartData': {'labels': [], 'values': []}, 'columns': []})

        processed_df, final_columns = process_report_data(df, 'main')

        for col in processed_df.select_dtypes(include=['datetime', 'datetimetz']).columns:
            processed_df[col] = processed_df[col].astype(str)
            
        if 'RecordTime' in processed_df.columns:
            processed_df = processed_df.drop(columns=['RecordTime'])
            if 'RecordTime' in final_columns: final_columns.remove('RecordTime')

        active_chems = [col for col in final_columns if col in CHEMICALS]
        # Calculate chart from all rows EXCEPT the last 'TOTAL WEIGHT' row
        chart_df = processed_df.iloc[:-1] if len(processed_df) > 1 else processed_df
        chart_values = [float(chart_df[col].sum()) for col in active_chems]

        return jsonify({
            'status': 'success',
            'tableData': processed_df.head(1000).fillna('-').to_dict(orient='records'),
            'chartData': {'labels': active_chems, 'values': chart_values},
            'columns': final_columns
        })
    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/excel', methods=['GET'])
def export_excel():
    try:
        start_date, end_date, report_type = request.args.get('startDate'), request.args.get('endDate'), request.args.get('reportType', 'main')
        query, params = build_query(start_date, end_date)
        conn = get_db_connection()
        df = pd.read_sql(query, conn, params=params)
        conn.close()

        # Check for empty data before doing Excel operations
        if df.empty:
            return "<script>alert('NO DATA AVAILABLE FOR THIS SELECTION'); window.history.back();</script>"

        processed_df, final_columns = process_report_data(df, report_type)
        if 'RecordTime' in processed_df.columns: processed_df = processed_df.drop(columns=['RecordTime'])
        
        for col in processed_df.select_dtypes(include=['datetime', 'datetimetz']).columns:
            processed_df[col] = processed_df[col].dt.tz_localize(None)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            processed_df.to_excel(writer, index=False, sheet_name='Report', columns=[c for c in final_columns if c != 'RecordTime'])
            worksheet = writer.sheets['Report']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column_letter].width = (max_length + 3) * 1.2

        output.seek(0)
        return send_file(output, download_name=f"SCADA_{report_type.capitalize()}_Report.xlsx", as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/pdf', methods=['GET'])
def export_pdf():
    try:
        start_date, end_date, report_type = request.args.get('startDate'), request.args.get('endDate'), request.args.get('reportType', 'main')
        query, params = build_query(start_date, end_date)
        conn = get_db_connection()
        df = pd.read_sql(query, conn, params=params)
        conn.close()
        
        # Check for empty data before doing PDF operations
        if df.empty:
            return "<script>alert('NO DATA AVAILABLE FOR THIS SELECTION'); window.history.back();</script>"
            
        processed_df, final_columns = process_report_data(df, report_type)
        if 'RecordTime' in processed_df.columns: processed_df = processed_df.drop(columns=['RecordTime'])
        
        pdf_df = processed_df[[c for c in final_columns if c != 'RecordTime']].fillna('-').astype(str)

        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=landscape(A0))
        elements = []
        styles = getSampleStyleSheet()
        elements.append(Paragraph(f"Chemical Consumption {report_type.capitalize()} Report", styles['Title']))
        
        data_list = [pdf_df.columns.tolist()] + pdf_df.values.tolist()
        t = Table(data_list)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'), 
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey) 
        ]))
        elements.append(t)
        doc.build(elements)
        output.seek(0)
        return send_file(output, download_name=f"SCADA_{report_type.capitalize()}_Report.pdf", as_attachment=True, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/export/chart-pdf', methods=['GET'])
def export_chart_pdf():
    try:
        start_date, end_date = request.args.get('startDate'), request.args.get('endDate')
        query, params = build_query(start_date, end_date)
        conn = get_db_connection()
        df = pd.read_sql(query, conn, params=params)
        conn.close()

        # Check for empty data before doing Chart operations
        if df.empty:
            return "<script>alert('NO DATA AVAILABLE FOR THIS SELECTION'); window.history.back();</script>"

        processed_df, final_columns = process_report_data(df, 'main')
        active_chems = [col for col in final_columns if col in CHEMICALS]
        chart_df = processed_df.iloc[:-1] if len(processed_df) > 1 else processed_df
        chart_values = [float(chart_df[col].sum()) for col in active_chems]

        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=landscape(A0))
        elements = []
        styles = getSampleStyleSheet()
        elements.append(Paragraph(f"Chemical Consumption Chart ({start_date} to {end_date})", styles['Title']))

        drawing = Drawing(2000, max(500, len(active_chems) * 30))
        bc = HorizontalBarChart()
        bc.x, bc.y = 150, 50
        bc.height, bc.width = max(400, len(active_chems) * 28), 1800
        bc.data = [chart_values]
        bc.categoryAxis.categoryNames = active_chems
        bc.valueAxis.valueMin = 0
        bc.bars[0].fillColor = colors.blue
        
        # Display Data Labels in KG 
        bc.barLabelFormat = '%.2f kg'
        bc.barLabels.nudge = 10
        bc.barLabels.boxAnchor = 'w' 
        
        drawing.add(bc)
        elements.append(drawing)
        doc.build(elements)
        output.seek(0)
        return send_file(output, download_name="SCADA_Chart.pdf", as_attachment=True, mimetype='application/pdf')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    print("🚀 SCADA Reporting Server running on network port 5555...")
    app.run(host='0.0.0.0', port=5555, debug=True)

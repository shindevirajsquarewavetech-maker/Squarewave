import os

# Define the target directory
TARGET_DIR = r"C:\Users\HP\OneDrive\Desktop\Tulsi SCADA NEW SOFTWARE upd"

# Ensure the directory exists
os.makedirs(TARGET_DIR, exist_ok=True)

# ---------------------------------------------------------
# 1. SERVER.PY CONTENT
# ---------------------------------------------------------
server_py_content = """import pyodbc
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
    'server': 'DESKTOP-42HFTV6\\\\SQLEXPRESS', 
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
"""

# ---------------------------------------------------------
# 2. INDEX.HTML CONTENT
# ---------------------------------------------------------
index_html_content = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>SCADA Chemical Usage Reporter</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="style.css">
</head>
<body class="flex h-screen overflow-hidden bg-slate-50">
    
    <aside class="w-72 bg-slate-900 text-slate-300 flex-shrink-0 hidden md:flex flex-col">
        <div class="p-6 border-b border-slate-800">
            <h1 class="text-2xl font-bold text-white tracking-wider flex items-center gap-2">
                <span class="text-emerald-500 text-3xl">☍</span> SCADA
            </h1>
        </div>
        <nav class="flex-grow p-4 space-y-4">
            <a href="#" class="block px-4 py-3 bg-slate-800 text-white rounded-xl shadow-md border-l-4 border-emerald-500 font-medium text-lg">
                📊 Main Dashboard
            </a>
            <hr class="border-slate-700">
            <h3 class="px-4 text-xs font-bold text-slate-500 uppercase tracking-wider">Specific Reports</h3>
            
            <button onclick="document.getElementById('daily-modal').classList.add('active')" class="w-full text-left block px-4 py-3 bg-slate-800 hover:bg-slate-700 text-white rounded-xl shadow-sm transition-colors font-medium text-lg border border-slate-700">
                📅 Generate Daily Report
            </button>
            <button onclick="document.getElementById('monthly-modal').classList.add('active')" class="w-full text-left block px-4 py-3 bg-slate-800 hover:bg-slate-700 text-white rounded-xl shadow-sm transition-colors font-medium text-lg border border-slate-700">
                📆 Generate Monthly Report
            </button>
        </nav>
    </aside>

    <main class="flex-grow flex flex-col h-screen overflow-y-auto">
        <header class="bg-white shadow-sm border-b border-slate-200 px-10 py-6 flex justify-between items-center sticky top-0 z-20">
            <h2 class="text-3xl font-bold text-slate-800">Chemical Consumption Report</h2>
            <div class="flex gap-4">
                <button onclick="downloadFile('excel')" class="bg-emerald-600 hover:bg-emerald-700 text-white px-6 py-3 rounded-lg font-bold shadow-sm transition-transform hover:-translate-y-0.5">📥 Excel Report</button>
                <button onclick="downloadFile('pdf')" class="bg-red-600 hover:bg-red-700 text-white px-6 py-3 rounded-lg font-bold shadow-sm transition-transform hover:-translate-y-0.5">📄 PDF Summary</button>
                <button onclick="downloadFile('chart-pdf')" class="bg-violet-600 hover:bg-violet-700 text-white px-6 py-3 rounded-lg font-bold shadow-sm transition-transform hover:-translate-y-0.5">📊 Chart PDF</button>
            </div>
        </header>

        <div class="p-6 space-y-8 max-w-full mx-auto w-full">
            <section class="bg-white rounded-xl shadow-sm border border-slate-200 p-6">
                <div class="grid grid-cols-1 md:grid-cols-3 gap-6 items-end">
                    <div>
                        <label class="block text-sm font-bold text-slate-600 uppercase tracking-wide mb-2">Start Date</label>
                        <input type="date" id="start-date" class="w-full p-4 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-lg font-medium text-slate-700">
                    </div>
                    <div>
                        <label class="block text-sm font-bold text-slate-600 uppercase tracking-wide mb-2">End Date</label>
                        <input type="date" id="end-date" class="w-full p-4 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-lg font-medium text-slate-700">
                    </div>
                    <div>
                        <button onclick="generateReport()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold p-4 rounded-lg shadow-sm transition-transform hover:-translate-y-0.5 flex justify-center items-center gap-2 text-lg">
                            <span id="gen-btn-text">Generate Dashboard</span>
                            <span id="gen-spinner" class="hidden animate-spin h-5 w-5 border-4 border-white border-t-transparent rounded-full"></span>
                        </button>
                    </div>
                </div>
            </section>

            <div id="error-container" class="hidden bg-red-50 border border-red-200 text-red-700 px-6 py-4 rounded-lg shadow-sm">
                <strong class="font-bold">Input Error:</strong> <span id="error-msg"></span>
            </div>

            <div id="results-area" class="space-y-8 hidden">
                <div class="bg-white p-6 rounded-xl shadow-sm border border-slate-200 cursor-pointer hover:border-blue-500 transition-colors" onclick="openChartModal()">
                    <div class="flex justify-between items-center mb-4">
                        <h3 class="text-xl font-bold text-slate-800">Top 10 Chemical Consumers (KG)</h3>
                        <span class="text-blue-600 font-medium text-sm bg-blue-50 px-3 py-1 rounded-md">⤢ Expand Chart</span>
                    </div>
                    <div style="height: 300px; width: 100%;"><canvas id="usageChartPreview"></canvas></div>
                </div>

                <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden cursor-pointer hover:border-blue-500 transition-colors">
                    <div class="flex justify-between items-center p-6 border-b border-slate-100 bg-slate-50" onclick="openTableModal()">
                        <h3 class="font-bold text-slate-800 text-xl">Batch Logs & Total Weight</h3>
                        <span class="text-blue-600 font-medium text-sm bg-blue-100 px-3 py-1 rounded-md">⤢ Expand Table</span>
                    </div>
                    <div class="overflow-x-auto" onclick="openTableModal()">
                        <table class="w-full text-left border-collapse">
                            <thead><tr id="table-header-preview" class="bg-slate-100 text-slate-600 text-sm uppercase tracking-wider"></tr></thead>
                            <tbody id="table-body-preview" class="divide-y divide-slate-100"></tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <div id="modal-overlay" class="modal-overlay">
        <div class="modal-content">
            <div class="modal-header">
                <h3 id="modal-title" class="text-xl font-bold text-slate-800">Detailed View</h3>
                <button onclick="closeModal('modal-overlay')" class="bg-red-500 hover:bg-red-600 text-white px-4 py-2 rounded-lg font-bold">✕ Close</button>
            </div>
            <div class="modal-body p-6 overflow-auto" id="modal-body"></div>
        </div>
    </div>

    <div id="daily-modal" class="modal-overlay">
        <div class="bg-white w-full max-w-md rounded-2xl shadow-2xl overflow-hidden">
            <div class="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                <h3 class="text-xl font-bold text-slate-800">📅 Daily Report</h3>
                <button onclick="closeModal('daily-modal')" class="text-slate-400 hover:text-red-500 font-bold text-xl">✕</button>
            </div>
            <div class="p-6 space-y-6">
                <div>
                    <label class="block text-sm font-bold text-slate-600 mb-2">Select Date</label>
                    <input type="date" id="daily-date" class="w-full p-3 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 text-lg">
                </div>
                <div class="flex gap-3">
                    <button onclick="downloadSpecial('daily', 'pdf')" class="flex-1 bg-red-600 hover:bg-red-700 text-white py-3 rounded-lg font-bold shadow-sm">📄 Get PDF</button>
                    <button onclick="downloadSpecial('daily', 'excel')" class="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-lg font-bold shadow-sm">📥 Get Excel</button>
                </div>
            </div>
        </div>
    </div>

    <div id="monthly-modal" class="modal-overlay">
        <div class="bg-white w-full max-w-md rounded-2xl shadow-2xl overflow-hidden">
            <div class="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
                <h3 class="text-xl font-bold text-slate-800">📆 Monthly Report</h3>
                <button onclick="closeModal('monthly-modal')" class="text-slate-400 hover:text-red-500 font-bold text-xl">✕</button>
            </div>
            <div class="p-6 space-y-6">
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm font-bold text-slate-600 mb-2">Start Year</label>
                        <input type="number" id="monthly-start-year" value="2024" class="w-full p-3 border border-slate-300 rounded-lg text-lg text-center">
                    </div>
                    <div>
                        <label class="block text-sm font-bold text-slate-600 mb-2">End Year</label>
                        <input type="number" id="monthly-end-year" value="2025" class="w-full p-3 border border-slate-300 rounded-lg text-lg text-center">
                    </div>
                </div>
                <div class="flex gap-3">
                    <button onclick="downloadSpecial('monthly', 'pdf')" class="flex-1 bg-red-600 hover:bg-red-700 text-white py-3 rounded-lg font-bold shadow-sm">📄 Get PDF</button>
                    <button onclick="downloadSpecial('monthly', 'excel')" class="flex-1 bg-emerald-600 hover:bg-emerald-700 text-white py-3 rounded-lg font-bold shadow-sm">📥 Get Excel</button>
                </div>
            </div>
        </div>
    </div>

    <script>
        // DYNAMIC HOST CONNECTION: This ensures the frontend connects to whatever IP address it was loaded from.
        const backendUrl = window.location.origin + "/api";
        let chartPreviewInstance = null; let chartFullInstance = null; let currentReportData = null;

        document.getElementById('start-date').valueAsDate = new Date();
        document.getElementById('end-date').valueAsDate = new Date();
        document.getElementById('daily-date').valueAsDate = new Date();
        document.getElementById('monthly-start-year').value = new Date().getFullYear();
        document.getElementById('monthly-end-year').value = new Date().getFullYear();

        function openChartModal() {
            if(!currentReportData) return;
            document.getElementById('modal-title').innerText = "Full Consumption Analysis";
            const body = document.getElementById('modal-body');
            body.innerHTML = '<div style="position:relative; width:100%; min-height:1200px;"><canvas id="usageChartFull"></canvas></div>';
            document.getElementById('modal-overlay').classList.add('active');

            setTimeout(() => {
                const ctx = document.getElementById('usageChartFull');
                if(chartFullInstance) chartFullInstance.destroy();
                chartFullInstance = new Chart(ctx, {
                    type: 'bar',
                    data: { labels: currentReportData.chartData.labels, datasets: [{ label: 'Total Usage (KG)', data: currentReportData.chartData.values, backgroundColor: '#3b82f6', borderRadius: 4 }] },
                    options: { indexAxis: 'y', responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
                });
            }, 100);
        }

        function openTableModal() {
            if(!currentReportData) return;
            document.getElementById('modal-title').innerText = "Complete Data & Totals";
            const body = document.getElementById('modal-body');
            let html = '<div class="overflow-x-auto"><table class="w-full text-left border-collapse"><thead><tr class="bg-slate-100 text-slate-600 uppercase text-sm">';
            currentReportData.columns.forEach(col => html += `<th class="p-3 border-b">${col}</th>`);
            html += '</tr></thead><tbody class="divide-y divide-slate-100">';
            
            currentReportData.tableData.forEach((row, i) => {
                const isTotalRow = (i === currentReportData.tableData.length - 1);
                html += `<tr class="${isTotalRow ? 'bg-blue-50 font-bold' : 'hover:bg-slate-50'}">`;
                currentReportData.columns.forEach(col => {
                    let val = row[col] !== null ? row[col] : '-';
                    html += `<td class="p-3 ${typeof val === 'number' ? 'text-right' : ''}">${val}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table></div>';
            body.innerHTML = html;
            document.getElementById('modal-overlay').classList.add('active');
        }

        function closeModal(id) { document.getElementById(id).classList.remove('active'); }

        async function generateReport() {
            const btnText = document.getElementById('gen-btn-text'); const spinner = document.getElementById('gen-spinner');
            const resultsArea = document.getElementById('results-area'); const errorContainer = document.getElementById('error-container');
            const startDate = document.getElementById('start-date').value; const endDate = document.getElementById('end-date').value;

            if(new Date(startDate) > new Date(endDate)) {
                document.getElementById('error-msg').innerText = "Start Date cannot be after End Date.";
                errorContainer.classList.remove('hidden'); return;
            }

            btnText.innerText = "Processing..."; spinner.classList.remove('hidden'); resultsArea.classList.add('hidden'); errorContainer.classList.add('hidden');

            try {
                const response = await fetch(`${backendUrl}/report`, {
                    method: 'POST', headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ startDate, endDate })
                });
                if (!response.ok) throw new Error("Server Error");
                const data = await response.json();
                if (data.status === 'error') throw new Error(data.message);

                currentReportData = data;
                renderPreviewChart(data.chartData); renderPreviewTable(data.columns, data.tableData);
                resultsArea.classList.remove('hidden'); 
            } catch (error) {
                document.getElementById('error-msg').innerText = error.message; errorContainer.classList.remove('hidden');
            } finally {
                btnText.innerText = "Generate Dashboard"; spinner.classList.add('hidden');
            }
        }

        function downloadFile(type) {
            const s = document.getElementById('start-date').value; const e = document.getElementById('end-date').value;
            window.location.href = `${backendUrl}/export/${type}?startDate=${s}&endDate=${e}&reportType=main`;
        }

        function downloadSpecial(reportType, fileType) {
            let s, e;
            if(reportType === 'daily') {
                s = e = document.getElementById('daily-date').value;
            } else {
                s = document.getElementById('monthly-start-year').value + "-01-01";
                e = document.getElementById('monthly-end-year').value + "-12-31";
            }
            window.location.href = `${backendUrl}/export/${fileType}?startDate=${s}&endDate=${e}&reportType=${reportType}`;
        }

        function renderPreviewChart(chartData) {
            const ctx = document.getElementById('usageChartPreview');
            if(chartPreviewInstance) chartPreviewInstance.destroy();
            const indices = chartData.values.map((val, i) => ({ val, label: chartData.labels[i] })).sort((a, b) => b.val - a.val).slice(0, 10);
            chartPreviewInstance = new Chart(ctx, {
                type: 'bar',
                data: { labels: indices.map(i => i.label), datasets: [{ label: 'Usage', data: indices.map(i => i.val), backgroundColor: '#3b82f6', borderRadius: 4 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true }, x: { grid: { display: false } } } }
            });
        }

        function renderPreviewTable(columns, rows) {
            const thead = document.getElementById('table-header-preview'); const tbody = document.getElementById('table-body-preview');
            thead.innerHTML = ''; tbody.innerHTML = '';
            columns.forEach(col => { const th = document.createElement('th'); th.className="p-4"; th.innerText = col; thead.appendChild(th); });
            
            // Show first 5 rows + the very last row (which contains the Totals)
            const previewRows = rows.slice(0, 4);
            if (rows.length > 4) previewRows.push(rows[rows.length - 1]);

            previewRows.forEach((row, index) => {
                const isTotalRow = (row[columns[0]] && row[columns[0]].toString().includes("TOTAL WEIGHT"));
                const tr = document.createElement('tr');
                tr.className = isTotalRow ? "bg-blue-50 font-bold border-t-2 border-blue-200" : "hover:bg-slate-50";
                
                columns.forEach(col => {
                    const td = document.createElement('td'); 
                    td.innerText = row[col] !== null ? row[col] : '-';
                    td.className = "p-4 " + (typeof row[col] === 'number' ? 'text-right' : '');
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }
    </script>
</body>
</html>"""

# ---------------------------------------------------------
# 3. REQUIREMENTS.TXT
# ---------------------------------------------------------
requirements_content = """flask
flask-cors
pyodbc
pandas
openpyxl
reportlab
"""

# ---------------------------------------------------------
# 4. RUN SERVER BAT
# ---------------------------------------------------------
bat_content = """@echo off
cd /d "%~dp0"
title SCADA Report Server
echo Starting Server...
python server.py
pause
"""

# ---------------------------------------------------------
# 5. STYLE.CSS
# ---------------------------------------------------------
css_content = """body { font-family: 'Inter', system-ui, sans-serif; }
.modal-overlay { position: fixed; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.7); z-index: 1000; display: none; justify-content: center; align-items: center; padding: 2rem; backdrop-filter: blur(4px); }
.modal-overlay.active { display: flex; }
.modal-content { background: white; width: 95%; height: 90%; border-radius: 1rem; display: flex; flex-direction: column; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.25); overflow: hidden; position: relative; }
.modal-header { padding: 1.5rem; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center; background: #f8fafc; }
.modal-body { flex-grow: 1; overflow: auto; background: #fff; }
"""

# Define files to create
files = {
    "server.py": server_py_content,
    "index.html": index_html_content,
    "requirements.txt": requirements_content,
    "run_server.bat": bat_content,
    "style.css": css_content
}

# Write files
for filename, content in files.items():
    filepath = os.path.join(TARGET_DIR, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Created/Updated: {filepath}")

print(f"\n✅ Software Successfully Compiled to: {TARGET_DIR}")
import base64
import datetime
import subprocess
import os
import dash
from dash import html, dcc, dash_table, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import pandas as pd
from data_processing import convert_file, get_df, apply_styling_to_excel, Employee

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

server = app.server

# Define the layout of the app
app.layout = dbc.Container([
    # Title
    dbc.Row(dbc.Col(html.Div(id='record-statistics'), width=12)),
    
    dbc.Row(dbc.Col(html.Label('Select Month and year', className='d-table', style={'margin': '40px'}), width=12)),
    dbc.Row(dbc.Col(dcc.DatePickerSingle(
        id='date-picker',
        placeholder='Select Month',
        date=datetime.date.today(),
    ), width=12)),
   
    # Upload component to select and upload PDF files
    dbc.Row(dbc.Col(html.Div([dcc.Upload(
        id='upload-pdf',
        children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files for the month')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        multiple=True
    )], className='columns'), width=12)),
    
    dbc.Row(dbc.Col(html.Button("Generate Dienst", id="btn-generate-dienst", style={'width': '100%', 'margin': '20px', 'margin-right': '0px'}), width=12)),
    dbc.Row(dbc.Col(html.Button("Download Excel", id="btn-download-excel", style={'width': '100%', 'margin': '20px', 'margin-right': '0px'}), width=12)),
    dcc.Download(id="download-excel"),
    
    # Progress bar and status
    dbc.Row(dbc.Col(dcc.Interval(id='interval-progress', interval=1000, n_intervals=0), width=12)),
    dbc.Row(dbc.Col(dbc.Progress(id='progress-bar', value=0, striped=True, animated=True, style={'margin': '20px'}), width=12)),
    dbc.Row(dbc.Col(html.Div(id='conversion-status'), width=12)),
        
    # DataFrame display area
    dbc.Row(dbc.Col(html.Div([
        dash_table.DataTable(
            id='data-table',
            style_table={'overflowX': 'scroll'},
            style_cell={
                'height': 'auto',
                'minWidth': '80px', 'width': '80px', 'maxWidth': '80px',
                'whiteSpace': 'normal',
                'textAlign': 'center',
                'border': '1px solid black'
            },
            style_header={
                'backgroundColor': 'rgb(230, 230, 230)',
                'fontWeight': 'bold',
                'border': '1px solid black'
            },
            style_data={
                'border': '1px solid black'
            },
            style_data_conditional=[
                {
                    'if': {
                        'filter_query': '{work_type_code} = "Free"',
                        'column_id': 'work_type_code'
                    },
                    'backgroundColor': '#3D9970',
                    'color': 'white'
                }
            ]
        )
    ]), width=12)),

], fluid=True)

@app.callback(
    Output("record-statistics", "children"),
    Output("data-table", "data"),
    Output("data-table", "columns"),
    Output("progress-bar", "value"),
    Output("conversion-status", "children"),
    Input("btn-generate-dienst", "n_clicks"),
    State('date-picker', 'date'),
    prevent_initial_call=True,
)
def generate_dienst(n_clicks, selected_date):
    df = None
    date_object = datetime.date.fromisoformat(selected_date)
    year = date_object.year
    month = date_object.month
    progress = 0
    status = []
    if os.path.exists(f"processed_data/{Employee}_{year}_{month}.parquet"):
        df = pd.read_parquet(f"processed_data/{Employee}_{year}_{month}.parquet")
    else:
        files = []
        total_files = len([f for f in os.listdir('uploaded_files') if f.endswith('.pdf')])
        for idx, filename in enumerate(os.listdir('uploaded_files')):
            if filename.endswith('.pdf'):
                with open(os.path.join('uploaded_files', filename), 'rb') as f:
                    file_content = f.read()
                    files.append(convert_file(filename, file_content))
                progress = int((idx + 1) / total_files * 100)
                status.append(f"Converted {filename} to Excel.")
        if files:
            files = [os.path.join('uploaded_files', f) for f in files]
            df = get_df(files, Employee, year, month)
            df.to_parquet(f"processed_data/{Employee}_{year}_{month}.parquet")
    if df is not None:
        df['date'] = df['date'].dt.date
        # Export to Excel with styling
        excel_path = f'processed_data/{Employee}_{year}_{month}.xlsx'
        df.to_excel(excel_path, index=False, engine='openpyxl')

        # Apply styling to the Excel file
        apply_styling_to_excel(excel_path)

        return html.H1(f'Dienst for Thomas Rager(TRG)'), df.to_dict("records"), [{'name': col, 'id': col} for col in df.columns], progress, html.Ul([html.Li(s) for s in status])
    else:
        return html.H1(f'Dienst for Thomas Rager(TRG)'), [], [], progress, html.Ul([html.Li(s) for s in status])

@app.callback(
    Output("download-excel", "data"),
    Input("btn-download-excel", "n_clicks"),
    State('date-picker', 'date'),
    prevent_initial_call=True,
)
def download_excel(n_clicks, selected_date):
    date_object = datetime.date.fromisoformat(selected_date)
    year = date_object.year
    month = date_object.month
    excel_path = f'processed_data/{Employee}_{year}_{month}.xlsx'
    
    # Apply styling to the Excel file
    apply_styling_to_excel(excel_path)
    
    return dcc.send_file(excel_path)

@app.callback(
    Output('upload-pdf', 'children'),
    Input('upload-pdf', 'contents'),
    State('upload-pdf', 'filename'),
    prevent_initial_call=True,
)
def store_files(contents, names):
    # Create the folder if it doesn't exist and clean everything in it
    upload_folder = 'uploaded_files'
    if not os.path.exists(upload_folder):
        os.makedirs(upload_folder)
    else:
        for filename in os.listdir(upload_folder):
            file_path = os.path.join(upload_folder, filename)
            if os.path.isfile(file_path):
                os.unlink(file_path)
    if contents is not None:
        for c, n in zip(contents, names):
            content_type, content_string = c.split(',')
            decoded = base64.b64decode(content_string)
            with open(os.path.join('uploaded_files', n), 'wb') as f:
                f.write(decoded)
        return html.Div(['Files uploaded successfully!'])
    return html.Div(['Drag and Drop or ', html.A('Select Files for the month')])

# Run the app
if __name__ == '__main__':
    if not os.path.exists('uploaded_files'):
        os.makedirs('uploaded_files')
    if not os.path.exists('processed_data'):
        os.makedirs('processed_data')
    subprocess.run(
        executable="playwright", args="install chromium", capture_output=True, check=True
    )
    app.run_server(debug=True)
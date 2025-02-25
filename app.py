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
    
    dbc.Row(dbc.Col(html.Button("Download Excel", id="btn-download-excel", style={'width': '100%', 'margin': '20px', 'margin-right': '0px'}), width=12)),
    dcc.Download(id="download-excel"),
        
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
    Output("record-statistics", "children"),
    Output("data-table", "data"),
    Output("data-table", "columns"),
    Input('upload-pdf', 'contents'),
    State('upload-pdf', 'filename'),
    State('date-picker', 'date')
)
def update_output(contents, names, selected_date):
    df = None
    date_object = datetime.date.fromisoformat(selected_date)
    year = date_object.year
    month = date_object.month
    if os.path.exists(f"processed_data/{Employee}_{year}_{month}.parquet"):
        df = pd.read_parquet(f"processed_data/{Employee}_{year}_{month}.parquet")
    elif contents is not None:
        files = []
        for c, n in zip(contents, names):
            content_type, content_string = c.split(',')
            decoded = base64.b64decode(content_string)
            files.append(convert_file(n, decoded))
        df = get_df(files, Employee, year, month)
        df.to_parquet(f"processed_data/{Employee}_{year}_{month}.parquet")
    if df is not None:
        df['date'] = df['date'].dt.date
        # Export to Excel with styling
        excel_path = f'processed_data/{Employee}_{year}_{month}.xlsx'
        df.to_excel(excel_path, index=False, engine='openpyxl')

        # Apply styling to the Excel file
        apply_styling_to_excel(excel_path)

        return html.H1(f'Dienst for Thomas Rager(TRG)'), df.to_dict("records"), [{'name': col, 'id': col} for col in df.columns]
    else:
        return html.H1(f'Dienst for Thomas Rager(TRG)'), [], []

# Run the app
if __name__ == '__main__':
    subprocess.run(
        executable="playwright", args="install chromium", capture_output=True, check=True
    )
    app.run_server(debug=True)
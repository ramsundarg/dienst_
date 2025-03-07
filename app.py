import base64
import datetime
import os
from time import sleep
import time
import dash
from dash import html, dcc, dash_table, State
from dash.dependencies import Input, Output
import dash_bootstrap_components as dbc
import pandas as pd
from data_processing import convert_file, get_df, apply_styling_to_excel

# Initialize the Dash app
from dash import DiskcacheManager
# Diskcache for non-production apps when developing locally
import diskcache
cache = diskcache.Cache("./cache")
background_callback_manager = DiskcacheManager(cache)

app = dash.Dash(__name__, background_callback_manager=background_callback_manager)

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
    
    # Input for employee code
    dbc.Row(dbc.Col(dcc.Input(
        id='employee-code',
        type='text',
        value='TRG',
        placeholder='Enter Employee Code',
        style={'width': '100%', 'margin': '20px'}
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
    dbc.Row(dbc.Col(        html.Div(
            [
                html.P(id="paragraph_id", children=["Button not clicked"]),
                html.Progress(id="progress_bar", value="0"),
            ]))),
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
    Input("btn-generate-dienst", "n_clicks"),
    State('date-picker', 'date'),
    State('employee-code', 'value'),
    running=[
        (Output("btn-generate-dienst", "disabled"), True, False),
        (Output("cancel_button_id", "disabled"), False, True),
        (
            Output("progress_bar", "style"),
            {"visibility": "visible"},
            {"visibility": "hidden"},
        ),
    ],
    #progress=[Output("progress_bar", "value"), Output("progress_bar", "max")],
    prevent_initial_call=True,
)
def generate_dienst(n_clicks,selected_date, employee_code):
    ctx = dash.callback_context

    if not ctx.triggered:
        print('No trigger')
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update, dash.no_update

    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    print(f'Triggered by: {trigger_id}')

    if trigger_id == 'btn-generate-dienst':
        df = None
        date_object = datetime.date.fromisoformat(selected_date)
        year = date_object.year
        month = date_object.month
        if os.path.exists(f"processed_data/{employee_code}_{year}_{month}.parquet"):
            df = pd.read_parquet(f"processed_data/{employee_code}_{year}_{month}.parquet")
        else:
            files = []
            total_files = len([f for f in os.listdir('uploaded_files') if f.endswith('.pdf')])
            for idx, filename in enumerate(os.listdir('uploaded_files')):
                #time.sleep(1)
                #set_progress((str(idx + 1), str(total_files)))
                if filename.endswith('.pdf'):
                    progress = int((idx + 1) / total_files * 100)
                    status=(f"Converted {filename} to Excel.")
                    with open(os.path.join('uploaded_files', filename), 'rb') as f:
                        file_content = f.read()
                        files.append(convert_file(os.path.join('uploaded_files', filename), file_content))
            if files:
                files = [os.path.join('uploaded_files',f) for f in os.listdir('uploaded_files') if f.endswith('.xlsx')]
                df = get_df(files, employee_code, year, month)
                df.to_parquet(f"processed_data/{employee_code}_{year}_{month}.parquet")
        if df is not None:
            df['date'] = df['date'].dt.date
            # Export to Excel with styling
            excel_path = f'processed_data/{employee_code}_{year}_{month}.xlsx'
            df.to_excel(excel_path, index=False, engine='openpyxl')

            # Apply styling to the Excel file
            apply_styling_to_excel(excel_path)
            print("Done in the background")
            

            return html.H1(f'Dienst for {employee_code}'), df.to_dict("records"), [{'name': col, 'id': col} for col in df.columns]
        else:
            return html.H1(f'Dienst for {employee_code}'), [], []


@app.callback(
    Output("download-excel", "data"),
    Input("btn-download-excel", "n_clicks"),
    State('date-picker', 'date'),
    State('employee-code', 'value'),
    prevent_initial_call=True,
)
def download_excel(n_clicks, selected_date, employee_code):
    date_object = datetime.date.fromisoformat(selected_date)
    year = date_object.year
    month = date_object.month
    excel_path = f'processed_data/{employee_code}_{year}_{month}.xlsx'
    
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
    
    app.run_server(debug=True)
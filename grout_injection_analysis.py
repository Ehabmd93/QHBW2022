import dash
from dash import dcc, html
import pandas as pd
import numpy as np
from datetime import timedelta
import os

# Your data processing functions
def calculate_averages(group, before_period, after_period):
    def get_non_zero_avg(data):
        numeric_data = data.apply(pd.to_numeric, errors='coerce')
        non_zero_flows = numeric_data[numeric_data['flow'] > 0]
        if len(non_zero_flows) == 0:
            return None, None, None
        avg_flow = non_zero_flows['flow'].mean()
        avg_eff_pressure = non_zero_flows['effPressure'].mean()
        avg_lugeon = non_zero_flows['Lugeon'].mean()
        return avg_flow, avg_eff_pressure, avg_lugeon
    
    while True:
        before_data = group[-before_period:]
        after_data = group[:after_period]
        avg_flow_before, avg_eff_pressure_before, avg_lugeon_before = get_non_zero_avg(before_data)
        avg_flow_after, avg_eff_pressure_after, avg_lugeon_after = get_non_zero_avg(after_data)
        
        if avg_flow_before is not None and avg_flow_after is not None:
            avg_flow = (avg_flow_before + avg_flow_after) / 2
            avg_eff_pressure = (avg_eff_pressure_before + avg_eff_pressure_after) / 2
            avg_lugeon = (avg_lugeon_before + avg_lugeon_after) / 2
            return avg_flow, avg_eff_pressure, avg_lugeon, before_period, after_period
        
        before_period += 5
        after_period += 5

def handle_last_mix(group):
    last_10_min = group[-10:]
    numeric_last_10_min = last_10_min.apply(pd.to_numeric, errors='coerce')
    min_flow = numeric_last_10_min['flow'].min()
    min_lugeon = numeric_last_10_min['Lugeon'].min()
    max_eff_pressure = numeric_last_10_min['effPressure'].max()
    last_marsh = group['vmarshGrout'].iloc[-1]
    return min_flow, min_lugeon, max_eff_pressure, last_marsh

def process_file(file_path):
    data = pd.read_excel(file_path)
    results = []
    unique_holes = data['holeNum'].unique()
    mix_counts = {'Water': 0, 'Mix A': 0, 'Mix B': 0, 'Mix C': 0, 'Mix D': 0}
    
    for hole in unique_holes:
        hole_data = data[data['holeNum'] == hole]
        hole_data = hole_data.sort_values(by='TIMESTAMP')
        stage_top = hole_data['stageTop'].iloc[0]
        stage_bottom = hole_data['stageBottom'].iloc[0]
    
        mix_types = hole_data['mixNum'].unique()
        for i, mix in enumerate(mix_types):
            mix_data = hole_data[hole_data['mixNum'] == mix]
            mix_start_time = mix_data['TIMESTAMP'].min()
            
            if i < len(mix_types) - 1:
                next_mix = mix_types[i + 1]
                next_mix_data = hole_data[hole_data['mixNum'] == next_mix]
                mix_end_time = next_mix_data['TIMESTAMP'].min()
            else:
                mix_end_time = mix_data['TIMESTAMP'].max() + timedelta(minutes=5)
            
            mix_duration = (mix_end_time - mix_start_time).total_seconds() / 60
            marsh_value = mix_data['vmarshGrout'].iloc[0]
            mix_volume = mix_data['volume'].iloc[-1] - mix_data['volume'].iloc[0]
            cumulative_volume = mix_data['volume'].iloc[-1]
            
            if i == len(mix_types) - 1:
                min_flow, min_lugeon, max_eff_pressure, last_marsh = handle_last_mix(mix_data)
                avg_flow = min_flow
                avg_lugeon = min_lugeon
                avg_eff_pressure = max_eff_pressure
                marsh_value = last_marsh
                extended_note = "Last mix: recorded minimum flow, minimum Lugeon, and maximum effective pressure over the last 10 minutes."
            else:
                avg_flow, avg_eff_pressure, avg_lugeon, before_period, after_period = calculate_averages(mix_data, 5, 5)
                extended_note = f"Extended period before: {before_period} minutes, after: {after_period} minutes"
            
            results.append([
                hole, 6, stage_top, stage_bottom, mix_start_time, mix_end_time, mix_duration,
                mix, marsh_value, mix_volume, cumulative_volume, avg_flow, avg_eff_pressure, avg_lugeon, extended_note
            ])
            
            if mix == 1:
                mix_counts['Water'] += 1
            elif mix == 2:
                mix_counts['Mix A'] += 1
            elif mix == 3:
                mix_counts['Mix B'] += 1
            elif mix == 4:
                mix_counts['Mix C'] += 1
            elif mix == 5:
                mix_counts['Mix D'] += 1
    
    columns = [
        'Hole ID', 'Stage', 'Stage Top', 'Stage Bottom', 'Time Start', 'Time Finish', 'Mix Duration (min)',
        'Mix', 'Marsh', 'Mix Volume', 'Cumulative Volume', 'Flow Avg (L/min)', 'Effective Pressure Avg (bar)',
        'GuL Avg', 'Extended Note'
    ]
    
    results_df = pd.DataFrame(results, columns=columns)
    
    output_dir = os.path.dirname(file_path)
    output_file = os.path.join(output_dir, 'grout_injection_summary.csv')
    
    try:
        results_df.to_csv(output_file, index=False)
        print(f"Processing complete. The summary has been saved to '{output_file}'.")
    except PermissionError as e:
        print(f"PermissionError: {e}. Please ensure the file is not open in another program.")
    
    mix_count_df = pd.DataFrame(list(mix_counts.items()), columns=['Mix Type', 'Count'])
    mix_count_file = os.path.join(output_dir, 'mix_count_summary.csv')
    mix_count_df.to_csv(mix_count_file, index=False)
    print(f"Mix count summary has been saved to '{mix_count_file}'.")

# Dash application setup
app = dash.Dash(__name__)

app.layout = html.Div(children=[
    html.H1(children='Grout Injection Analysis'),

    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files')
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
    ),

    html.Div(id='output-data-upload'),
])

def parse_contents(contents, filename):
    content_type, content_string = contents.split(',')

    decoded = base64.b64decode(content_string)
    file_path = f'/tmp/{filename}'
    with open(file_path, 'wb') as f:
        f.write(decoded)
    
    process_file(file_path)
    return html.Div([
        html.H5(filename),
        html.H6('Processing complete.'),
    ])

@app.callback(Output('output-data-upload', 'children'),
              [Input('upload-data', 'contents')],
              [State('upload-data', 'filename')])
def update_output(list_of_contents, list_of_names):
    if list_of_contents is not None:
        children = [
            parse_contents(c, n) for c, n in zip(list_of_contents, list_of_names)
        ]
        return children

server = app.server

if __name__ == '__main__':
    app.run_server(debug=True)

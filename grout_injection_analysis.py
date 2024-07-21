import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import os
import re
from dash import Dash, dcc, html, Input, Output, State
from dash.exceptions import PreventUpdate
from datetime import timedelta

# Function to extract Hole ID, Order, and Stage from the filename
def extract_file_details(file_name):
    try:
        pattern = r'([PS]\d{4})_(S\d+)'  # double backslashes for escaping in strings
        match = re.search(pattern, file_name)
        if match:
            hole_id, stage = match.groups()
            order = 'P' if hole_id.startswith('P') else 'S'
            return hole_id, order, stage
    except Exception as e:
        print(f"Error extracting file details from {file_name}: {e}")
    return None, None, None

# Function to clean and preprocess the data
def clean_data(data):
    try:
        if 'TOA5' in data.columns:
            data.columns = data.iloc[0]
            data = data[1:]
            data.reset_index(drop=True, inplace=True)
            data.drop(index=[0, 2, 3], inplace=True)
        else:
            data.drop(index=[0, 2, 3], inplace=True)
        data['TIMESTAMP'] = pd.to_datetime(data['TIMESTAMP'])
        return data
    except Exception as e:
        print(f"Error cleaning data: {e}")
        return None

# Function to identify Marsh change events
def identify_marsh_changes(data):
    try:
        marsh_changes = data[data['vmarshGrout'].diff() != 0]  # Identify rows where vmarshGrout changes
        marsh_changes = marsh_changes[['TIMESTAMP', 'mixNum', 'vmarshGrout', 'flow']]
        return marsh_changes
    except Exception as e:
        print(f"Error identifying Marsh changes: {e}")
        return None

# Function to generate the interactive graph and return the data
def generate_interactive_graph(file_path):
    try:
        print(f"Processing file: {file_path}")
        data = pd.read_excel(file_path, sheet_name='Raw Data')
        data = clean_data(data)
        if data is None:
            return None, None

        marsh_changes = identify_marsh_changes(data)
        if marsh_changes is None:
            return None, None

        hole_id = data['holeNum'].iloc[0] if 'holeNum' in data.columns else 'Unknown'
        start_time = data['TIMESTAMP'].min()
        data['ElapsedMinutes'] = (data['TIMESTAMP'] - start_time).dt.total_seconds() / 60

        fig = go.Figure()

        # Plot each mix's flow rate in a different color
        mix_colors = {2: 'blue', 3: 'cyan', 4: 'magenta', 5: 'orange'}
        mix_labels = {2: 'A', 3: 'B', 4: 'C', 5: 'D'}
        for mix, color in mix_colors.items():
            mix_data = data[data['mixNum'] == mix]
            if not mix_data.empty:
                add_trace(fig, mix_data, f'Flow Rate Mix {mix_labels[mix]}', 'flow', color)
                add_trace(fig, mix_data, f'Effective Pressure Mix {mix_labels[mix]}', 'effPressure', 'green', yaxis='y2')

        # Add vertical lines for each mix change event
        mix_changes = data[data['mixNum'].diff().abs() > 0]
        for _, row in mix_changes.iterrows():
            mix_change_min = (row['TIMESTAMP'] - start_time).total_seconds() / 60
            fig.add_vline(x=mix_change_min, line=dict(color='red', width=2, dash='dash'))

        # Add small dots and annotations for Marsh changes
        for _, row in marsh_changes.iterrows():
            marsh_change_min = (row['TIMESTAMP'] - start_time).total_seconds() / 60
            fig.add_trace(go.Scatter(
                x=[marsh_change_min],
                y=[row['flow']],
                mode='markers+text',
                marker=dict(color='black', size=8),
                text=[f"{row['vmarshGrout']:.1f}"],
                textposition='top center'
            ))

        fig.update_layout(title='Flow Rate and Effective Pressure vs Time for Mixes', 
                          xaxis=dict(title='Time Elapsed (Minutes)'),
                          yaxis=dict(title='Flow Rate (L/min)', side='left'), 
                          yaxis2=dict(title='Effective Pressure (bar)', overlaying='y', side='right'),
                          hovermode='x unified')
        return fig, data
    except Exception as e:
        print(f"Error generating interactive graph for file {file_path}: {e}")
        return None, None

# Helper function to add a trace to the figure
def add_trace(fig, data, name, y_col, color, yaxis='y'):
    fig.add_trace(go.Scatter(
        x=data['ElapsedMinutes'],
        y=data[y_col],
        mode='lines+markers',
        name=name,
        line=dict(color=color),
        text=[f"Flow: {row['flow']:.1f}, Eff Pressure: {row['effPressure']:.1f}, Lugeon: {row['Lugeon']:.1f}" 
              for _, row in data.iterrows()],
        yaxis=yaxis
    ))

# Function to generate a scatter plot
def generate_scatter_plot(data):
    try:
        fig = px.scatter(data, x='flow', y='effPressure', color='mixNum', 
                         labels={'flow': 'Flow Rate', 'effPressure': 'Effective Pressure'},
                         title='Flow Rate vs Effective Pressure')
        return fig
    except Exception as e:
        print(f"Error generating scatter plot: {e}")
        return None

# Function to generate a histogram
def generate_histogram(data, column):
    try:
        fig = px.histogram(data, x=column, nbins=20, title=f'Distribution of {column}')
        return fig
    except Exception as e:
        print(f"Error generating histogram: {e}")
        return None

# Function to generate a box plot
def generate_box_plot(data, column):
    try:
        fig = px.box(data, y=column, title=f'Box Plot of {column}')
        return fig
    except Exception as e:
        print(f"Error generating box plot: {e}")
        return None

# Function to process all files and return a list of file details
def get_all_files_details(input_directory):
    try:
        if not os.path.exists(input_directory):
            print(f"The directory {input_directory} does not exist.")
            return []
        files_details = []
        for filename in os.listdir(input_directory):
            if filename.endswith('.xlsx'):
                hole_id, order, stage = extract_file_details(filename)
                if hole_id and order and stage:
                    files_details.append({'filename': filename, 'hole_id': hole_id, 'order': order, 'stage': stage})
        return files_details
    except Exception as e:
        print(f"Error processing files in directory {input_directory}: {e}")
        return []

# Dash App
app = Dash(__name__)
input_directory = r'C:\\Payhton\\Panda\\Panda\\Loop\\Excel Raw'
output_directory = r'C:\\Payhton\\Panda\\Panda\\Loop\\Output-Graphs'
files_details = get_all_files_details(input_directory)
hole_ids = sorted(list(set([f['hole_id'] for f in files_details])))
orders = sorted(list(set([f['order'] for f in files_details])))
stages = sorted(list(set([f['stage'] for f in files_details])))

# Debugging: Print the extracted file details
print("Files details:", files_details)

# Debugging: Print the unique values for dropdowns
print("Hole IDs:", hole_ids)
print("Orders:", orders)
print("Stages:", stages)

app.layout = html.Div([
    html.H1("Grout Injection Data Analysis"),
    html.Div([
        html.Label('Hole ID'),
        dcc.Dropdown(id='hole-id-dropdown', options=[{'label': hole_id, 'value': hole_id} for hole_id in hole_ids], 
                     value=hole_ids[0] if hole_ids else None, placeholder="Select a Hole ID")
    ], style={'width': '30%', 'display': 'inline-block'}),
    html.Div([
        html.Label('Order'),
        dcc.Dropdown(id='order-dropdown', options=[{'label': order, 'value': order} for order in orders], 
                     value=orders[0] if orders else None, placeholder="Select an Order")
    ], style={'width': '30%', 'display': 'inline-block'}),
    html.Div([
        html.Label('Stage'),
        dcc.Dropdown(id='stage-dropdown', options=[{'label': stage, 'value': stage} for stage in stages], 
                     value=stages[0] if stages else None, placeholder="Select a Stage")
    ], style={'width': '30%', 'display': 'inline-block'}),
    html.Div([
        html.Label('View Type'),
        dcc.RadioItems(id='view-type', options=[
            {'label': 'Time Series', 'value': 'time_series'},
            {'label': 'Scatter Plot', 'value': 'scatter'},
            {'label': 'Histogram', 'value': 'histogram'},
            {'label': 'Box Plot', 'value': 'box'}
        ], value='time_series')
    ], style={'width': '30%', 'display': 'inline-block'}),
    dcc.Graph(id='interactive-graph'),
    html.Div(id='clicked-point-data', style={'marginTop': 20})
])

@app.callback(
    Output('interactive-graph', 'figure'),
    [Input('hole-id-dropdown', 'value'),
     Input('order-dropdown', 'value'),
     Input('stage-dropdown', 'value'),
     Input('view-type', 'value')]
)
def update_graph(selected_hole_id, selected_order, selected_stage, view_type):
    if not selected_hole_id or not selected_order or not selected_stage:
        print("Preventing update due to missing input")
        raise PreventUpdate

    # Debugging: Print selected values
    print(f"Selected Hole ID: {selected_hole_id}, Order: {selected_order}, Stage: {selected_stage}, View Type: {view_type}")

    # Find the corresponding file
    for file_detail in files_details:
        if file_detail['hole_id'] == selected_hole_id and file_detail['order'] == selected_order and file_detail['stage'] == selected_stage:
            file_path = os.path.join(input_directory, file_detail['filename'])
            print(f"Selected file: {file_path}")  # Debugging statement to ensure correct file path
            fig, data = generate_interactive_graph(file_path)
            if data is None:
                print("No data found, preventing update")
                raise PreventUpdate
            if view_type == 'time_series':
                fig, _ = generate_interactive_graph(file_path)
            elif view_type == 'scatter':
                fig = generate_scatter_plot(data)
            elif view_type == 'histogram':
                fig = generate_histogram(data, 'flow')
            elif view_type == 'box':
                fig = generate_box_plot(data, 'flow')
            return fig

    print("No matching file found, preventing update")
    raise PreventUpdate

@app.callback(
    Output('clicked-point-data', 'children'),
    Input('interactive-graph', 'clickData')
)
def display_click_data(clickData):
    if clickData is None:
        raise PreventUpdate

    point_data = clickData['points'][0]
    custom_data = point_data['text']
    return f"Selected Point Data: {custom_data}"

if __name__ == '__main__':
    app.run_server(debug=True)

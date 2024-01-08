# To run this script:

# Ensure you have Dash installed. If not, install it using pip install dash dash-renderer dash-html-components dash-core-components plotly.
# Run the script. It will start a local web server.
# Open your web browser and go to http://127.0.0.1:8050/.
# The application will display in your browser with a text input field and dynamic output that updates with every keystroke.

# The script uses Dash, a high-level framework for building web applications in Python, which is ideal for this task.
# The use of dcc and html modules from the Dash library helps in creating interactive components and HTML layout respectively.
# The callback function is used to update the output text dynamically based on the input, showcasing an interactive feature of Dash apps.
# The run_server method starts the app, making it accessible in a web browser.


import dash
import dash_core_components as dcc
import dash_html_components as html
import dash_table
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px

# Load the spreadsheet data
file_path = r'C:\Users\Madhous\Documents\GitHub\EUD_AM\Asset Management\Dash Web App\EUC_Perth_Assets.xlsx'
df_items = pd.read_excel(file_path, sheet_name='4.2 Items')
df_timestamps = pd.read_excel(file_path, sheet_name='4.2 Timestamps')
df_all_sans = pd.read_excel(file_path, sheet_name='All SANs')

# Initialize the Dash app
app = dash.Dash(__name__)

app.layout = html.Div([
    html.Div([
        html.H1("Spreadsheet Data and Graphs", style={'textAlign': 'center', 'color': '#0074D9'}),
        html.Div([
            dash_table.DataTable(
                id='table-items',
                columns=[{"name": i, "id": i} for i in df_items.columns],
                data=df_items.to_dict('records'),
                style_table={'overflowX': 'scroll'},
                style_cell={'minWidth': '180px', 'width': '180px', 'maxWidth': '180px', 'overflow': 'hidden',
                            'textOverflow': 'ellipsis'},
                style_header={
                    'backgroundColor': '#0074D9',
                    'fontWeight': 'bold',
                    'color': 'white'
                },
                style_data_conditional=[
                    {
                        'if': {'row_index': 'odd'},
                        'backgroundColor': 'rgb(248, 248, 248)'
                    }
                ]
            ),
        ], style={'width': '50%', 'display': 'inline-block', 'padding': '20px'}),

        html.Div([
            dcc.Graph(id='graph-items'),
        ], style={'width': '50%', 'display': 'inline-block', 'padding': '20px'}),

        # Repeat the same structure for other tables and graphs
    ]),
    # Additional styling or components here
], style={'fontFamily': 'Arial, sans-serif'})

@app.callback(
    [Output('graph-items', 'figure')],
    [Input('table-items', 'data')]
)
def update_graphs(data_items):
    df_items = pd.DataFrame(data_items)
    fig_items = px.line(df_items, x='YourXColumn', y='YourYColumn')  # Customize as needed
    fig_items.update_layout(
        plot_bgcolor='white',
        paper_bgcolor='white',
        font=dict(color='black', size=12),
        hovermode='closest',
        xaxis=dict(showgrid=False, zeroline=False, showline=False, showticklabels=True),
        yaxis=dict(showgrid=False, zeroline=False, showline=False, showticklabels=True)
    )
    return [fig_items]

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)

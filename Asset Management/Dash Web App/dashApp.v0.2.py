# This script creates a web application with a text input and a dynamic text output. The output text changes as you type in the input field.

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
df = pd.read_excel(r'C:\Users\Madhous\Documents\GitHub\EUD_AM\Asset Management\Dash Web App\EUC_Perth_Assets.xlsx', sheet_name='4.2 Items')

# Initialize the Dash app
app = dash.Dash(__name__)

# Define the app layout
app.layout = html.Div([
    html.H1("Spreadsheet Data and Graphs"),
    dash_table.DataTable(
        id='table',
        columns=[{"name": i, "id": i} for i in df.columns],
        data=df.to_dict('records'),
    ),
    html.Div([
        dcc.Graph(id='graph-1'),
        dcc.Graph(id='graph-2')
        # Add more graphs as needed
    ])
])

# Define callbacks for updating graphs
@app.callback(
    [Output('graph-1', 'figure'),
     Output('graph-2', 'figure')],
    [Input('table', 'data')]
)
def update_graphs(table_data):
    dff = pd.DataFrame(table_data)
    
    # Create plotly graphs
    fig1 = px.line(dff, x='YourXColumn', y='YourYColumn')  # Modify as needed
    fig2 = px.bar(dff, x='YourXColumn', y='YourYColumn')   # Modify as needed
    
    return fig1, fig2

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)

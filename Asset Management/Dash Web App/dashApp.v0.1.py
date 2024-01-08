import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output

# Initialize the Dash app
app = dash.Dash(__name__)

# Define the app layout
app.layout = html.Div([
    html.H1("Welcome to the Dash App"),
    dcc.Input(id='input-text', value='Initial Value', type='text'),
    html.Div(id='output-text')
])

# Define callback to update output
@app.callback(
    Output('output-text', 'children'),
    [Input('input-text', 'value')]
)
def update_output(input_value):
    return f'You entered: {input_value}'

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
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
from dash.dependencies import Output
import pandas as pd
import plotly.express as px
import numpy as np

# Load the spreadsheet data
file_path = r'C:\Users\Madhous\Documents\GitHub\EUD_AM\Asset Management\Dash Web App\EUC_Perth_Assets.xlsx'
df_items = pd.read_excel(file_path, sheet_name='4.2 Items')

# Initialize the Dash app
app = dash.Dash(__name__)

# Prepare data for boxplot
item_column = df_items.columns[0]  # Assuming first column contains items
volume_column = 'LastCount'  # Replace with the actual column name for volume
df_items[volume_column] = 1  # Assuming each row is one item
item_counts = df_items[item_column].value_counts()

# Create a DataFrame suitable for boxplot
boxplot_data = pd.DataFrame({
    item_column: np.repeat(item_counts.index, item_counts.values),
    volume_column: np.repeat(item_counts.values, item_counts.values)
})

# Define the app layout
app.layout = html.Div([
    html.H1("Item Volume Boxplot", style={'textAlign': 'center'}),
    dcc.Graph(id='boxplot-graph')
])

@app.callback(
    Output('boxplot-graph', 'figure'),
    []
)
def update_graph():
    print("update_graph function called")  # Check if the function is called

    max_volume = boxplot_data[volume_column].max() + 20
    fig = px.box(boxplot_data, x=item_column, y=volume_column, color=item_column)

    # Printing data details for debugging
    print(f"Max volume: {max_volume}")
    print(f"Unique items: {boxplot_data[item_column].unique()}")

    # Customizing the plot
    fig.update_traces(boxmean='sd')
    fig.update_layout(
        yaxis_range=[0, max_volume],
        plot_bgcolor='white',
        showlegend=False
    )

    # Adding the number of items on the boxes
    for i, item in enumerate(boxplot_data[item_column].unique()):
        count = item_counts[item]
        print(f"Item: {item}, Count: {count}, Annotation Position: {i}, {max_volume-2}")  # Debugging annotation positions

        fig.add_annotation(
            x=i, y=max_volume-2,
            text=str(count),
            showarrow=False,
            font=dict(size=10, color='black')
        )
    return fig

# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)


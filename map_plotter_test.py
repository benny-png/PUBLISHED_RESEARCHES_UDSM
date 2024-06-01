# Import necessary libraries
import pandas as pd
import json
import plotly.express as px

# Load your local GeoJSON file
with open('TZA_adm1_mkoaTZ.geojson') as f:
    geojson_data = json.load(f)

# Load the CSV file containing food production data
food_production_data = pd.read_csv('food_production_data.csv')

# Convert 'TOTAL PRODUCTION' column to numeric, handling errors as NaN
food_production_data['TOTAL PRODUCTION'] = pd.to_numeric(food_production_data['TOTAL PRODUCTION'], errors='coerce')

# Replace invalid or missing values (represented by "-") with NaN
food_production_data.replace("-", float('nan'), inplace=True)

# Fill NaN values in 'TOTAL PRODUCTION' column with 0
food_production_data['TOTAL PRODUCTION'].fillna(0, inplace=True)

# Generate color scale based on total production
color_scale = px.colors.sequential.Greens

# Create choropleth map
fig = px.choropleth(
    food_production_data,
    geojson=geojson_data,  # Load the GeoJSON data
    locations="REGION",    # Column in CSV file with region names
    featureidkey="properties.NAME_1",  # Key in GeoJSON file that matches the locations
    color="TOTAL PRODUCTION",  # Column in CSV file for color coding
    color_continuous_scale=color_scale,  # Color scale based on total production
    range_color=(0, food_production_data['TOTAL PRODUCTION'].max()),  # Range of color scale
    scope="africa",  # Setting scope to Africa
    title="Food Crop Production Distribution in Tanzania-mainland Regions"  # Map title
)

# Update the layout of the map
fig.update_layout(
    title_x=0.5,  # Set the title's horizontal position to the center
    title_y=0.9,  # Set the title's vertical position to the top
    title_font_size=24,  # Set the font size of the title
    title_font_family="Arial",  # Set the font family of the title
    annotations=[
        dict(
            x=0.5,
            y=-0.05,
            showarrow=False,
            text="IMAGE CREATED USING A COMPUTER ALGORITHM FROM OFFICIAL DATA (OF 2019)",
            xref="paper",
            yref="paper",
            font=dict(
                family="Arial",
                size=12,
                color="black"
            )
        )
    ]
)

# Update the layout of the map to center on Tanzania and zoom in
fig.update_geos(
    fitbounds="locations",
    visible=False
)

fig.show()

# Import necessary libraries
import pandas as pd
import geopandas as gpd
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

# Create a GeoDataFrame from the GeoJSON data
gdf = gpd.GeoDataFrame.from_features(geojson_data['features'])
gdf = gdf.merge(food_production_data, left_on='NAME_1', right_on='REGION')

# Function to find neighboring regions
def find_neighbors(gdf):
    neighbors = {}
    for index, row in gdf.iterrows():
        region = row['NAME_1']
        region_neighbors = gdf[gdf.geometry.touches(row['geometry'])]['NAME_1'].tolist()
        neighbors[region] = region_neighbors
    return neighbors

# Find neighbors for each region
neighbors = find_neighbors(gdf)

# Initialize variables
used_regions = set()
networks = []
CONSTANT = 20771405 / 3  # Example constant value = TOTAL_PRODUCTION/NO_OF_MANUFACTURING_INDUSTRIES

while len(used_regions) < len(gdf):
    # Initialize the starting region
    remaining_gdf = gdf[~gdf['NAME_1'].isin(used_regions)]
    if remaining_gdf.empty:
        break
    start_region = remaining_gdf.loc[remaining_gdf['TOTAL PRODUCTION'].idxmax()]
    selected_regions = [start_region['NAME_1']]
    total_production = start_region['TOTAL PRODUCTION']

    # Iterate to find bounding regions
    while total_production < CONSTANT:
        current_neighbors = []
        for region in selected_regions:
            current_neighbors.extend(neighbors[region])
        current_neighbors = list(set(current_neighbors) - set(selected_regions) - used_regions)
        
        if not current_neighbors:
            break
        
        next_region = gdf.loc[gdf['NAME_1'].isin(current_neighbors)].sort_values('TOTAL PRODUCTION', ascending=False).iloc[0]
        selected_regions.append(next_region['NAME_1'])
        total_production += next_region['TOTAL PRODUCTION']

    used_regions.update(selected_regions)
    networks.append((selected_regions, total_production))

# Print each network of neighboring regions and their total production
for i, (network, production) in enumerate(networks):
    print(f"Network {i + 1}: {network}, Total Production: {production}")

# Create a DataFrame to store the network information
network_data = []
for i, (network, production) in enumerate(networks):
    for region in network:
        network_data.append({'REGION': region, 'NETWORK': i})

network_df = pd.DataFrame(network_data)

# Merge the network data with the original food production data
food_production_data = food_production_data.merge(network_df, on='REGION', how='left')

# Generate color scale based on networks
network_colors = px.colors.qualitative.Plotly

# Create choropleth map
fig = px.choropleth(
    food_production_data,
    geojson=geojson_data,
    locations="REGION",
    featureidkey="properties.NAME_1",
    color="NETWORK",
    color_discrete_sequence=network_colors,
    scope="africa",
    title="FOOD CROP PRODUCTION DISTRIBUTION IN TANZANIA-MAINLAND REGIONS"
)

# Update the layout of the map
fig.update_layout(
    title_x=0.5,
    title_y=0.9,
    title_font_size=24,
    title_font_family="Arial",
    annotations=[
        dict(
            x=0.5,
            y=-0.05,
            showarrow=False,
            text="GENERATED IMAGE BY A COMPUTER ALGORITHM USING OFFICIAL DATA (OF 2019)",
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

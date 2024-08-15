import geopandas as gpd
import pandas as pd


pd.set_option('display.max_columns', 20)
file_path = './data/CA_Counties'
gis_data = gpd.read_file(file_path)

print(gis_data)

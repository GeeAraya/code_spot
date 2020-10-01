#python code to read a geojson file and convert it to an ee geometry obeject for AOI filters
#this code can be added to a python script as a utility
workingpath = 'G:\My Drive\YourJSONFile.json'

json_geo_df = geopandas.read_file(workingpath)
json_geo_js = json_geo_df.to_json()
json_geo_js_dict = json.loads(json_geo_js)
for feature in json_geo_js_dict['features']:
    json_geo = feature['geometry']['coordinates']
area = ee.Geometry.Polygon (json_geo)
print(area)

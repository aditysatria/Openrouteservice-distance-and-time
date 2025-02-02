import folium
import pandas as pd
import openrouteservice as ors
from openrouteservice.distance_matrix import distance_matrix

api_key = '#your api key here'
client = ors.Client(key=api_key)

file_path = 'D:\\KP\\koordinat.xlsx'  # Path to the uploaded file
spbe_data = pd.read_excel(file_path, sheet_name='Koordinat ori') # source sheet name
terminal_data = pd.read_excel(file_path, sheet_name='Koordinat dest') # destination sheet name

# 3. Menyiapkan koordinat
spbe_locations = spbe_data[['longitude', 'latitude']].values.tolist()
terminal_locations = terminal_data[['longitude', 'latitude']].values.tolist()


spbe_coords = [[lon, lat] for lat, lon in spbe_locations]
terminal_coords = [[lon, lat] for lat, lon in terminal_locations]

# 4. Menghitung jarak dari setiap SPBE ke setiap terminal
results = []
for spbe_idx, spbe in enumerate(spbe_coords):
    for terminal_idx, terminal in enumerate(terminal_coords):
        try:
            # Hitung jarak
            route = client.directions(
                coordinates=[spbe, terminal],
                profile='driving-car',
                format='json'
            )
            distance = route['routes'][0]['summary']['distance']  # in meter
            duration = route['routes'][0]['summary']['duration']  # in second

            # Simpan hasil
            results.append({
                "SPBE": spbe_data.iloc[spbe_idx]['SPBE'],
                "Terminal": f"Terminal {terminal_idx + 1}",
                "Distance (m)": distance,
                "Duration (s)": duration
            })

        except Exception as e:
            print(f"Error processing SPBE {spbe_idx+1} to Terminal {terminal_idx+1}: {e}")

# 5. convert result to DataFrame
df_results = pd.DataFrame(results)

# 6. Save to excel
df_results.to_excel("D:\\KP\\spbe_to_terminal_distances.xlsx", index=False)
print("Hasil disimpan dalam spbe_to_terminal_distances.xlsx")

# 7. visual to map
mymap = folium.Map(location=spbe_coords[0], zoom_start=1)

# add destination point to map
for idx, spbe in enumerate(spbe_coords):
    folium.Marker(location=[spbe[1], spbe[0]], icon=folium.Icon(color='blue'),
                  tooltip=f"SPBE {idx + 1}").add_to(mymap)

# add souce point to map
for idx, terminal in enumerate(terminal_coords):
    folium.Marker(location=[terminal[1], terminal[0]], icon=folium.Icon(color='red'),
                  tooltip=f"Terminal {idx + 1}").add_to(mymap)

# save the map to HTML
mymap.save('D:\\KP\\spbe_to_terminal_map.html') #your file destination
print("Peta disimpan sebagai 'spbe_to_terminal_map.html'")

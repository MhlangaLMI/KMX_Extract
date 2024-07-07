import zipfile
import xml.etree.ElementTree as ET
import requests
import math

# Define your Google Maps Elevation API key
API_KEY = "YOUR_API_KEY_HERE"

def get_elevation(lat, lng):
    url = f"https://maps.googleapis.com/maps/api/elevation/json?locations={lat},{lng}&key={API_KEY}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        if data['status'] == 'OK':
            return data['results'][0]['elevation']
        else:
            raise ValueError(f"Error fetching elevation data: {data['status']}")
    else:
        raise ConnectionError(f"Error connecting to the API: {response.status_code}")

def geo_distance(lat1, lon1, lat2, lon2):
    earth_radius_km = 6371
    lat1 = math.radians(lat1)
    lon1 = math.radians(lon1)
    lat2 = math.radians(lat2)
    lon2 = math.radians(lon2)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    distance_km = earth_radius_km * c
    return distance_km * 1000  # Convert to meters

def decimal_to_dms(decimal_degrees, direction):
    degrees = int(decimal_degrees)
    minutes = int((decimal_degrees - degrees) * 60)
    seconds = (decimal_degrees - degrees - minutes / 60) * 3600
    if direction == "lat":
        hemisphere = "N" if decimal_degrees >= 0 else "S"
    elif direction == "lon":
        hemisphere = "E" if decimal_degrees >= 0 else "W"
    else:
        hemisphere = ""
    return f"{abs(degrees)}° {abs(minutes)}' {abs(seconds):.2f}\" {hemisphere}"

def extract_coordinates(kmz_file_path):
    with zipfile.ZipFile(kmz_file_path, 'r') as kmz:
        with kmz.open('doc.kml', 'r') as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()
            namespace = {"kml": "http://www.opengis.net/kml/2.2"}

            coordinates_list = []

            for placemark in root.findall(".//kml:Placemark", namespaces=namespace):
                for coords in placemark.findall(".//kml:coordinates", namespaces=namespace):
                    coord_text = coords.text.strip()
                    coord_pairs = coord_text.split()
                    previous_lat = None
                    previous_lon = None
                    cumulative_distance = 0
                    for pair in coord_pairs:
                        lon, lat, _ = map(float, pair.split(","))
                        elevation = get_elevation(lat, lon)
                        dms_lat = decimal_to_dms(lat, "lat")
                        dms_lon = decimal_to_dms(lon, "lon")
                        distance = 0 if previous_lat is None else geo_distance(previous_lat, previous_lon, lat, lon)
                        cumulative_distance += distance
                        coordinates_list.append((lat, lon, elevation, dms_lat, dms_lon, distance, cumulative_distance))
                        previous_lat = lat
                        previous_lon = lon
            
            return coordinates_list

if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python extract_coordinates.py <path_to_kmz_file>")
        sys.exit(1)
    
    kmz_file_path = sys.argv[1]
    try:
        coordinates = extract_coordinates(kmz_file_path)
        for coord in coordinates:
            print(f"{coord[0]},{coord[1]},{coord[2]},{coord[3]},{coord[4]},{coord[5]},{coord[6]}")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

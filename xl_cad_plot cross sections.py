import xlwings as xw
import pyautocad
import pythoncom
import win32com.client
from collections import defaultdict



tab = "S1"

material_colors = {
    'Glacial Till - Clay': 5,
    'Highly Weathered Rock - CFG': 40,
    'Partially Weathered Rock - CFG': 8,
    'Competent Rock - CFG': 1,
    'Topsoil': 23,
    'No Recovery': 0
}


def extract_data_from_excel(file_path, tab):
    try:
        WB = xw.Book(file_path)
        WS = WB.sheets[tab]
        last_row = WS.range("A2").end("down").row
        data = WS.range(f"A2:D{last_row}").value
        return data
    except Exception as e:
        print(f"Error extracting data from Excel: {e}")
        return None

def get_autocad():
    try:
        acad = pyautocad.Autocad(create_if_not_exists=True)
        return acad
    except Exception as e:
        print(f"Error getting AutoCAD instance: {e}")
        return None

def variants(object):
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, (object))

def cad_plot_borehole_data(data):
    acad = get_autocad()
    if acad is None:
        return

    try:
        # Get the model space
        acad_model = acad.ActiveDocument.ModelSpace
        for rl_top, rl_bottom, chainage, legend in data:
            height = rl_top - rl_bottom
            color = material_colors.get(legend, 0)  # Default to gray for unknown legends

            # Create a new points list for each iteration
            points = [
                (chainage, rl_bottom, 0),
                (chainage + 2.5, rl_bottom, 0),
                (chainage + 2.5, rl_bottom + height, 0),
                (chainage, rl_bottom + height, 0),
                (chainage, rl_bottom, 0)
            ]

            # Flatten the points list and convert to aDouble
            flattened_points = [coord for point in points for coord in point]
            #print(flattened_points)
            points_array = pyautocad.aDouble(*flattened_points)  # Unpack the flattened points

            # Add the polyline to the model space
            polyline = acad_model.AddPolyline(points_array)
            polyline.color = color

    except Exception as e:
        print(f"Error plotting borehole data in AutoCAD: {e}")




if __name__ == "__main__":
    file_path = r"C:\Users\921722\Royal HaskoningDHV\Project-PC5301-Scapa-Deep-Water-Quay - Team\WIP\03_Geotechnical\01-Ground Modelling\101_Land-area\00 Prelim Model\Stratigraphy Summary_R1.xlsx"
    extracted_data = extract_data_from_excel(file_path, tab)

    if extracted_data:
        cad_plot_borehole_data(extracted_data)
        #print(extracted_data)
        # Initialize a dictionary to store summarized data
        summary = defaultdict(lambda: {'top': float('-inf'), 'bottom': float('inf'), 'material_type': None})
        # Iterate through the data
        for item in extracted_data:
            top_elevation, bottom_elevation, chainage, material_type = item
            if top_elevation > summary[chainage]['top']:
                summary[chainage]['top'] = top_elevation
            if bottom_elevation < summary[chainage]['bottom']:
                summary[chainage]['bottom'] = bottom_elevation
            summary[chainage]['material_type'] = material_type
        # Output the summarized data
        for chainage, info in sorted(summary.items()):
            print(
                f"Chainage: {chainage}, Material Type: {info['material_type']}, Highest Top Elevation: {info['top']}, Bottommost Bottom Elevation: {info['bottom']}")

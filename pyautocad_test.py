import xlwings as xw
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np
import tkinter as tk

tab = "S1"

material_colors = {
    'Glacial Till - Clay': 'blue',
    'Highly Weathered Rock - CFG': 'orange',
    'Partially Weathered Rock - CFG': 'gray',
    'Competent Rock - CFG': 'red',
    'Topsoil': 'brown',
    'No Recovery': 'black'
}

def extract_data_from_excel(file_path, tab):
    WB = xw.Book(file_path)
    WS = WB.sheets[tab]
    last_row = WS.range("A2").end("down").row
    data = WS.range(f"A2:D{last_row}").value
    return data

def plot_borehole_data(data, tab):
    fig, ax = plt.subplots(figsize=(10, 6))
    handles = []
    labels = []
    for rl_top, rl_bottom, chainage, legend in data:
        height = rl_top - rl_bottom
        color = material_colors.get(legend, 'gray')  # Default to gray if material is not in the dictionary
        rect = patches.Rectangle((chainage, rl_bottom), 7.5, height, linewidth=1, edgecolor='black', facecolor=color, label=legend)
        if legend not in labels:
            handles.append(rect)
            labels.append(legend)
        ax.add_patch(rect)

    # Set labels and title
    ax.set_xlabel('Chainage')
    ax.set_ylabel('RL (mCD)')
    plt_title = tab.replace("S", "")
    ax.set_title(f"Section {plt_title}")

    # Set axis limits
    min_chainage = min(data, key=lambda x: x[2])[2]
    max_chainage = max(data, key=lambda x: x[2])[2]
    min_rl = min(data, key=lambda x: x[1])[1]
    max_rl = max(data, key=lambda x: x[0])[0]
    ax.set_xlim(min_chainage - 5, max_chainage + 100)  # Add some padding to the sides
    ax.set_ylim(min_rl - 5, max_rl + 10)  # Add some padding to the top and bottom
    ax.minorticks_on()
    ax.grid(True, which='both', linestyle=':', linewidth='0.5')

    # Initialize a list to store selected points
    selected_points = []

    def on_click(event):
        if event.inaxes is not None:  # Check if the click is within the plot area
            selected_points.append((event.xdata, event.ydata))
            ax.plot(event.xdata, event.ydata, 'ro')  # Plot selected point
            fig.canvas.draw()

    fig.canvas.mpl_connect('button_press_event', on_click)

    plt.show()

    return selected_points

def select_points_for_material(data, material):
    print(f"Select points for {material} polygon:")
    return plot_borehole_data(data, tab)

def extract_unique_coordinates(data):
    unique_coordinates = []
    # Iterate through each data point
    for rl_top, rl_bottom, chainage, _ in data:
        # Add top and bottom coordinates as tuples
        top_coordinates = (chainage, rl_top)
        bottom_coordinates = (chainage, rl_bottom)
        # Add unique top coordinates
        if top_coordinates not in unique_coordinates:
            unique_coordinates.append(top_coordinates)
        # Add unique bottom coordinates
        if bottom_coordinates not in unique_coordinates:
            unique_coordinates.append(bottom_coordinates)
    return unique_coordinates

def find_closest_point(selected_point, unique_coordinates):
    # Convert the selected point to a numpy array for vectorized computation
    selected_point = np.array(selected_point)
    # Calculate the distances between the selected point and all unique coordinates
    distances = np.linalg.norm(unique_coordinates - selected_point, axis=1)
    # Find the index of the closest point
    closest_index = np.argmin(distances)
    # Return the closest point
    return unique_coordinates[closest_index]


def replace_selected_points(selected_points, unique_coordinates):
    replaced_points = []
    # Iterate through each selected point
    for point in selected_points:
        # Find the closest point in the unique coordinate list
        closest_point = find_closest_point(point, unique_coordinates)
        # Add the closest point to the replaced points list
        replaced_points.append(closest_point)
    return replaced_points

def main():
    file_path = r"C:\Users\921722\Royal HaskoningDHV\Project-PC5301-Scapa-Deep-Water-Quay - Team\WIP\03_Geotechnical\01-Ground Modelling\101_Land-area\00 Prelim Model\Stratigraphy Summary_R1.xlsx"
    extracted_data = extract_data_from_excel(file_path, tab)
    print(extracted_data)
    if extracted_data:
        # Extract unique coordinates from the raw data
        unique_coordinates = extract_unique_coordinates(extracted_data)
        for material in material_colors:
            selected_points = select_points_for_material(extracted_data, material)
            # Replace selected points with closest points from unique coordinate list
            replaced_points = replace_selected_points(selected_points, unique_coordinates)
            print(f"Replaced points for {material}:", replaced_points)

if __name__ == "__main__":
    main()

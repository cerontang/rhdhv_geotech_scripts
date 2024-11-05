import xlwings as xw
import matplotlib.pyplot as plt
import matplotlib.patches as patches

tab = "Sec_plot"

material_colors = {
    # SILT
    "Firm SILT": (34, 139, 34),  # Forest Green
    "Soft SILT": (0, 255, 0),  # Lime Green

    # CLAY
    "Firm CLAY": (255, 165, 0),  # Orange
    "Soft CLAY": (255, 255, 0),  # Yellow


    # PEAT
    "Firm PEAT": (128, 0, 128),  # Purple
    "Soft PEAT": (219, 112, 147),  # Pale Violet Red

    # SAND and GRAVEL
    "SAND and GRAVEL": (80, 80, 80)  # Charcoal
}




# Function to convert RGB tuple to hex color
def rgb_to_hex(rgb):
    return '#%02x%02x%02x' % rgb


def extract_data_from_excel(file_path, tab):
    WB = xw.Book(file_path)
    WS = WB.sheets[tab]
    last_row = WS.range("A2").end("down").row
    # Include column G (Location ID) in the extraction
    data = WS.range(f"A2:G{last_row}").value
    return data


def plot_borehole_data(data, tab):
    # Extract x and y values, material types, and location IDs
    x_values = [chainage for _, _, chainage, _, _, _, _ in data]
    y_values_top = [rl_top for rl_top, _, _, _, _, _, _ in data]
    y_values_bottom = [rl_bottom for _, rl_bottom, _, _, _, _, _ in data]
    materials = [material for _, _, _, material, _, _, _ in data]
    location_ids = [location_id for _, _, _, _, _, _, location_id in data]

    # Calculate min and max values for x and y
    min_x = min(x_values)
    max_x = max(x_values)
    min_y = min(min(y_values_top), min(y_values_bottom))
    max_y = max(max(y_values_top), max(y_values_bottom))

    # Create a dictionary to store the bottommost y-value for each borehole
    bottommost_y = {}

    # Plotting
    plt.figure(figsize=(14, 8))
    handles = {}
    for rl_top, rl_bottom, chainage, material, _, _, location_id in data:
        height = rl_top - rl_bottom
        color = rgb_to_hex(
            material_colors.get(material, (169, 169, 169)))  # Default to Dark Grey if material is not in the dictionary
        rect = patches.Rectangle((chainage, rl_bottom), 7.5, height, linewidth=1, edgecolor='black', facecolor=color)
        if material not in handles:
            handles[material] = rect
        plt.gca().add_patch(rect)

        # Update the bottommost y-value for each borehole
        if chainage not in bottommost_y or rl_bottom < bottommost_y[chainage]:
            bottommost_y[chainage] = rl_bottom

    # Add label for each borehole location ID, 1 meter below the bottommost layer
    for chainage, bottom in bottommost_y.items():
        plt.text(chainage + 7.5, bottom - 1, location_ids[x_values.index(chainage)], fontsize='small', ha='center')

    # Sort handles and labels based on the order of material_colors dictionary
    sorted_legends = {legend: handles[legend] for legend in material_colors if legend in handles}
    sorted_handles = list(sorted_legends.values())
    sorted_labels = list(sorted_legends.keys())

    # Set labels and title
    plt.xlabel('Chainage')
    plt.ylabel('RL (mOD)')
    plt.title(f"PC6313 Ennis South Flood Relief Scheme")

    # Set axis limits
    plt.xlim(min_x - 100, max_x + 100)
    plt.ylim(min_y - 1, max_y + 1)
    plt.minorticks_on()
    plt.grid(True, which='both', linestyle=':', linewidth='0.5')

    # Set legend properties
    plt.legend(sorted_handles, sorted_labels, loc='center left', bbox_to_anchor=(1, 0.5), borderaxespad=0.,
               fontsize='xx-small', title='Material', title_fontsize='medium')

    # Adjust layout to make room for the legend
    plt.tight_layout(rect=[0, 0, 1, 1])

    # Save the plot
    plt.savefig(
        f"C:/Users/921722/Royal HaskoningDHV/P-PC6313-Ennis-South-Flood-Relief-Schem - Team/WIP/7. Tasks/2. RHDHV Ground Intepretation/1. OG Outputs/Cross Section {tab}.png")
    plt.show()



file_path = r"C:\Users\921722\Royal HaskoningDHV\P-PC6313-Ennis-South-Flood-Relief-Schem - Team\WIP\7. Tasks\2. RHDHV Ground Intepretation\GI Data Summary_20240723.xlsm"
extracted_data = extract_data_from_excel(file_path, tab)
print(extracted_data)
plot_borehole_data(extracted_data, tab)

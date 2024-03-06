# Попробуем загрузить CSV файл и обработать его
import pathlib

import numpy as np
import pandas as pd

# try:
# Загрузка данных из CSV файла
workdir = str(pathlib.Path(pathlib.Path.cwd().parent,'data','dataset3','data_sulfate.csv'))

import matplotlib.pyplot as plt

# Since we cannot directly extract data from the image, we assume that the user wants to create a similar plot
# based on the previously formatted data we have saved.
# We will create a horizontal bar plot with error bars representing the ranges.
def parse_range(value):
    # Remove brackets and split the string into lower and upper values
    lower, upper = str(value).strip('[]').split(';')
    # Convert to float and handle the comma as a decimal separator
    lower = float(lower.replace(',', '.'))
    upper = float(upper.replace(',', '.'))
    return lower, upper


# Read the formatted data
new_formatted_data = pd.read_csv(workdir, sep=';')
unique_categories = new_formatted_data['Category'].unique()

# Create a dictionary to hold the plots for each category
category_plots = {}

# Generate a plot for each unique category
for category in unique_categories:
    if category =='Sulfide':
        # Filter the data for the current category
        category_data = new_formatted_data[new_formatted_data['Category'] == category]

        # Extract lower, upper bounds and calculate the means for the current category plot
        category_lower_bounds = []
        category_upper_bounds = []
        category_means = []
        for value in category_data['103 δ34SVCDT']:
            lower, upper = parse_range(value)
            category_lower_bounds.append(lower)
            category_upper_bounds.append(upper)
            category_means.append((upper + lower) / 2)

        # Calculate errors from the mean to the lower and upper bounds for current category data
        category_errors_lower = [mean - lower for mean, lower in zip(category_means, category_lower_bounds)]
        category_errors_upper = [upper - mean for mean, upper in zip(category_means, category_upper_bounds)]

        # Subcategory numbers for the y-axis for current category data
        category_subcategory_numbers = list(range(1, len(category_means) + 1))

        # Create the plot for the current category
        fig, ax = plt.subplots(figsize=(10, 6))

        # Plot horizontal bars
        ax.errorbar(category_means, category_subcategory_numbers, xerr=[category_errors_lower, category_errors_upper],  linestyle='None', ecolor='blue', capsize=5)

        # Set labels
        ax.set_xlabel('δ 34S, x 10³')
        # ax.set_ylabel('Subcategory')

        # Set the title to the category name
        ax.set_title(f'Variations of sulfur in {category}')

        # Add subcategory names as y-tick labels without the category name
        ax.set_yticks(category_subcategory_numbers)
        ax.set_yticklabels(category_data['Subcategory'], rotation='horizontal')

        # Invert the y-axis to match the given image
        ax.invert_yaxis()

        # Show the plot with a tight layout to ensure the labels fit well
        plt.tight_layout()
        plt.savefig(str(pathlib.Path(pathlib.Path.cwd(),'doc','img','sulfide_intervals.png')))
        plt.show()

sulfide_data = new_formatted_data[new_formatted_data['Category'] == 'Sulfide']

 # We will reconstruct the histogram to reflect the number of intersecting intervals on the y-axis.

# To do this, we'll create a list of all edges of the intervals
edges = []
for value in sulfide_data['103 δ34SVCDT']:
    lower, upper = parse_range(value)
    edges.append(lower)
    edges.append(upper)

# Sort the edges
edges = sorted(list(set(edges)))


# Create an array to hold the count of overlaps at each point
overlaps = np.zeros(len(edges) - 1)
print('overlaps', overlaps)

# Count the overlaps for each interval defined by these edges
bins = []
for i in range(len(overlaps)):
    for value in sulfide_data['103 δ34SVCDT']:
        lower, upper = parse_range(value)
        print(value, lower, upper)
        # If the edge is within an interval, increment the count
        if lower <= edges[i] < upper or lower < edges[i+1] <= upper:
            overlaps[i] += 1


# Define the bin centers as the midpoint between edges
bin_centers = (np.array(edges[:-1]) + np.array(edges[1:])) / 2



# Plotting the step plot
plt.figure(figsize=(8, 6))
plt.step(bin_centers, overlaps, where='mid', color='blue', linewidth=3)

max_overlap = np.max(overlaps)
print(max_overlap)
# Find the indices where the overlap is equal to the max overlap
mode_indices = np.where(overlaps == max_overlap)[0]

# Plotting the step plot with mode(s) highlighted
plt.figure(figsize=(8, 6))
plt.step(bin_centers, overlaps, where='mid', color='blue', linewidth=3)

# Loop through the mode indices to highlight the modes
for mode_index in mode_indices:
    plt.bar(bin_centers[mode_index], overlaps[mode_index], width=edges[mode_index+1] - edges[mode_index], color='red', alpha=0.3, edgecolor='red')

print(bin_centers)
print(bin_centers[mode_indices[0]], max_overlap)
# Annotations for the number of modes
plt.text(bin_centers[mode_indices[0]], max_overlap, f'Количество мод: {int(max_overlap)}              ', color='red', ha='right')

# Labels and title
plt.xlabel('δ 34S, x 10³')
plt.ylabel('μ_i')
# plt.title('Пересечение интервалов сульфидов в единицах 10³ δ34SVCDT')

# Show the plot

# Show the plot
plt.savefig(str(pathlib.Path(pathlib.Path.cwd(),'doc','img','sulfide_mode.png')))

plt.show()

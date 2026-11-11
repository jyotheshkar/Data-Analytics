"""This script performs data analysis on all the csv files."""
"""This script also downloads the graph into a folder named Images and creates one if there isn't a folder named images."""
"""This script also downloads a report into a folder named reports and creates one if there isn't a folder named reports."""


# All the necassary libraries are imported 
import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import zscore
import os
import xlsxwriter

def weekly_summary(data):
    """Calculate weekly counts without changing the caller's data or opening files."""
    if 'Created Date' not in data.columns:
        raise ValueError('CSV must contain a Created Date column')
    if data.empty:
        raise ValueError('CSV must contain at least one booking')
    dates = pd.to_datetime(data['Created Date'], format='%d/%m/%Y', errors='raise')
    if dates.isna().any():
        raise ValueError('Created Date cannot be empty')
    first_monday = dates.min() - pd.Timedelta(days=dates.min().weekday())
    week_numbers = ((dates - first_monday).dt.days // 7) + 1
    weeks = range(1, int(week_numbers.max()) + 1)
    counts = week_numbers.value_counts().reindex(weeks, fill_value=0).sort_index()
    result = pd.DataFrame({'Sequential Week Number': list(weeks), 'Number of Bookings': counts.to_numpy()})
    mean = result['Number of Bookings'].mean()
    std = result['Number of Bookings'].std(ddof=0)
    result['z-score'] = 0.0 if std == 0 else ((result['Number of Bookings'] - mean) / std).round(4)
    return result, mean, std, mean + 2 * std

# Define the main function for booking activity analysis and plotting
def plot_booking_activity_analysis_and_display_table(file_path, skiprows=0):
    df = pd.read_csv(file_path, skiprows=skiprows)
    booking_counts_complete, mean_bookings, std_bookings, threshold_value = weekly_summary(df)
    all_weeks = booking_counts_complete
    # Sets the plot style and creates a figure and axes for plotting
    plt.style.use('classic')
    fig, ax = plt.subplots(figsize=(16, 12))
    # Sets the figure and axes the background color
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    # Plots the number of bookings per week as bars
    bars = ax.bar(booking_counts_complete['Sequential Week Number'], booking_counts_complete['Number of Bookings'], color='skyblue', alpha=0.7, width=1)
    # Annotates each bar with the height (number of bookings)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{int(height)}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    # 3 points vertical offest
                    xytext=(0, 3),  
                    textcoords="offset points",
                    ha='center', va='bottom')
    # highlights weeks with bookings above the threshold value
    for i, row in booking_counts_complete.iterrows():
        color = 'red' if row['Number of Bookings'] > threshold_value else 'black'
        ax.scatter(row['Sequential Week Number'], row['Number of Bookings'], color=color)
    # Sets the x and y labels
    ax.set_xlabel('Week No.')
    ax.set_ylabel('Number of Bookings')
    # Sets the title for the plot
    filename = os.path.basename(file_path)
    ax.set_title(f'Booking Count by Week ({filename})')
    # Add gridlines for easier readability
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    # Adjusts x-axis ticks which is based on the number of weeks
    total_weeks = len(all_weeks)
    if total_weeks > 50:
        ax.set_xticks(range(1, total_weeks + 1, 5))
    else:
        ax.set_xticks(range(1, total_weeks + 1))
    ax.set_xlim(0.5, total_weeks + 0.5)
    # Draws a horizontal line at the threshold value
    ax.axhline(y=threshold_value, color='darkorange', linestyle='--', linewidth=2)
    # Adds the text label for the threshold value
    ax.text(len(all_weeks) + 0.5, threshold_value, f"Threshold: {threshold_value:.2f}", va='center', ha='left', color='white', backgroundcolor='black', bbox=dict(facecolor='black', edgecolor='orange', boxstyle="round,pad=0.3"))
    plt.tight_layout()
    # Saves the plot image to a file
    base_directory = os.path.dirname(file_path)
    filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
    image_directory = os.path.join(base_directory, "Images")
    os.makedirs(image_directory, exist_ok=True)
    image_file_path = os.path.join(image_directory, f"{filename_without_extension}.png")
    plt.savefig(image_file_path, dpi=300, bbox_inches='tight')
    print(f"Plot image saved at: {image_file_path}")
    plt.show()
    # Creates an excel file with booking data and analysis
    excel_filename = f"datareport_{filename_without_extension}.xlsx"
    excel_folder = os.path.join(base_directory, "reports")
    os.makedirs(excel_folder, exist_ok=True)
    excel_filepath = os.path.join(excel_folder, excel_filename)
    workbook = xlsxwriter.Workbook(excel_filepath)
    worksheet = workbook.add_worksheet()
    # Defines formats for the excel sheet headers and data and for styling them to
    heading_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black', 'border': 1})
    data_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000'})
    headings = ['Sequential Week Number', 'Number of Bookings', 'z-score']
    # Writes the column headings
    worksheet.write_row('A1', headings, heading_format)
    # Writes the data rows, applying background color for z-scores above the value 2
    for row, data in enumerate(booking_counts_complete.values, start=1):
        for col, value in enumerate(data):
            format_to_use = data_format
            if col == 2 and value > 2:
                format_to_use = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            worksheet.write(row, col, value, format_to_use)
    # Writes the summary data below the main table
    summary_row_start = len(booking_counts_complete) + 3
    summary_data = [("Threshold", threshold_value), ("Mean", mean_bookings), ("Standard Deviation", std_bookings)]
    summary_heading_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black', 'border': 1})
    summary_data_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000'})
    for i, (label, value) in enumerate(summary_data):
        worksheet.write(summary_row_start + i, 0, label, summary_heading_format)
        worksheet.write(summary_row_start + i, 1, value, summary_data_format)
    # Closes the workbook to save the excel file
    workbook.close()
    print(f"Excel file saved successfully: {excel_filepath}")

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Create weekly booking charts and an Excel report.')
    parser.add_argument('csv_file', help='CSV with a Created Date column in DD/MM/YYYY format')
    parser.add_argument('--skiprows', type=int, default=0, help='Use 1 for historical files with an extra metadata row')
    args = parser.parse_args()
    plot_booking_activity_analysis_and_display_table(args.csv_file, skiprows=args.skiprows)

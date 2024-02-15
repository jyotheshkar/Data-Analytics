import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import zscore
import numpy as np
import os
import xlsxwriter  # Importing xlsxwriter for Excel file writing

def plot_combined_booking_activity_analysis_and_display_table(file_path):
    # To read the CSV file and to skip the first row if it is necessary
    df = pd.read_csv(file_path, skiprows=1)
    
    # To convert the 'Created Date' column to datetime format
    df['Created Date'] = pd.to_datetime(df['Created Date'], format='%d/%m/%Y')
    
    # To find the first and last dates in the data
    start_date = df['Created Date'].min()
    end_date = df['Created Date'].max()
    
    # To calculate the number of days and subtract to get to the nearest Monday 
    days_to_monday = (start_date.weekday() - 0) % 7
    
    # Adjusting the start_date to the nearest Monday
    first_monday = start_date - pd.Timedelta(days=days_to_monday)
    
    # To Calculate sequential week numbers which starts from the first monday
    df['Sequential Week Number'] = ((df['Created Date'] - first_monday).dt.days // 7) + 1

    # Creating a dataframe that wil list all sequential week numbers
    all_weeks = pd.DataFrame({'Sequential Week Number': range(1, ((end_date - first_monday).days // 7) + 2)})
    
    # To group the data by week and then count the number of bookings per week 
    booking_counts = df.groupby('Sequential Week Number').size().reset_index(name='Number of Bookings')
    
    # To merge the list of all the weeks with the booking counts and also to include the weeks with no booking counts
    booking_counts_complete = pd.merge(all_weeks, booking_counts, on='Sequential Week Number', how='left').fillna(0)
    
    # Calculate the z-scores for the booking counts to identify the outliers
    booking_counts_complete['z-score'] = zscore(booking_counts_complete['Number of Bookings'], ddof=0)
    
    # Round the z-scores to 4 decimal places 
    booking_counts_complete['z-score'] = booking_counts_complete['z-score'].round(4)
    
    # start plotting
    plt.style.use('classic')
    fig, ax = plt.subplots(figsize=(16, 12))
    fig.patch.set_facecolor('white')  # changes the background to white
    ax.set_facecolor('white')  # changes the plot background to white

    # To plot the number of bookings per week in a bar chart 
    ax.bar(booking_counts_complete['Sequential Week Number'], booking_counts_complete['Number of Bookings'], color='skyblue', alpha=0.7, width=1)

    # for threshold value calculate the mean and standard deviation 
    mean_bookings = booking_counts_complete['Number of Bookings'].mean()
    std_bookings = booking_counts_complete['Number of Bookings'].std(ddof=0)
    threshold_value = mean_bookings + std_bookings

    # To determine the offset for text annotations or tooltips based on the range of bookings each week 
    offset = max(booking_counts_complete['Number of Bookings']) * 0.02

    # To annotate each bar with the total number of bookings and also highlighting the ones above the threshold in red
    for i, row in booking_counts_complete.iterrows():
        color = 'red' if row['Number of Bookings'] > threshold_value else 'black'
        ax.scatter(row['Sequential Week Number'], row['Number of Bookings'], color=color, zorder=5)
        ax.text(row['Sequential Week Number'], row['Number of Bookings'] + offset, f"{int(row['Number of Bookings'])}", 
                ha='center', va='center', color='white', fontsize=8, 
                bbox=dict(facecolor=color, edgecolor='none', boxstyle="round,pad=0.3"))
    
    # Adjusts the x-axis tick labels based on the total number of weeks
    if len(all_weeks) > 50:
        tick_spacing = 5
        ticks = range(0, len(all_weeks) + tick_spacing, tick_spacing)
        ax.set_xticks(ticks)
        ax.set_xticklabels(ticks, rotation=0)  # Display x-axis labels upright
    else:
        ax.set_xticks(range(1, booking_counts_complete['Sequential Week Number'].max() + 1))
        ax.set_xticklabels(range(1, booking_counts_complete['Sequential Week Number'].max() + 1), rotation=0)

    # Sets the axis title and labels
    ax.set_xlabel('Week No.')
    ax.set_ylabel('Number of Bookings')
    filename = os.path.basename(file_path)
    ax.set_title(f'Booking Count by Week ({filename})')

    # Adjusts the grid, ticks and limits for better visualization
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_xlim(0.5, len(all_weeks) + 0.5)
    ax.set_ylim(0, max(booking_counts_complete['Number of Bookings']) + (5 * offset))

    # To draw the threshold line in dotted format and annotate it with the threshold value
    threshold_line = ax.axhline(y=threshold_value, color='darkorange', linestyle='--', linewidth=2)
    ax.text(len(all_weeks) + 0.5, threshold_value, f"Threshold: {threshold_value:.2f}", 
            va='center', ha='left', 
            bbox=dict(facecolor='black', edgecolor='orange', boxstyle="round,pad=0.3", linewidth=2), color='orange')

    plt.tight_layout()  # Layout Adjustments for clean visual
    plt.show()  # Displays the plots

    # Save the table data to an Excel file
    excel_filename = f"datareport_{os.path.splitext(os.path.basename(file_path))[0]}.xlsx"
    excel_folder = os.path.join(os.path.dirname(file_path), "reports")
    os.makedirs(excel_folder, exist_ok=True)  # Create 'reports' folder if not exists
    excel_filepath = os.path.join(excel_folder, excel_filename)
    workbook = xlsxwriter.Workbook(excel_filepath)
    worksheet = workbook.add_worksheet()

    # Add heading formatting
    heading_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black', 'border': 1})
    data_format = workbook.add_format({'border': 1})
    red_bg_format = workbook.add_format({'bg_color': 'red', 'border': 1})

    # Write table data
    headings = ['Sequential Week Number', 'Number of Bookings', 'z-score']
    for col, heading in enumerate(headings):
        worksheet.write(0, col, heading, heading_format)

    for row, data in enumerate(booking_counts_complete.values):
        for col, value in enumerate(data):
            if headings[col] == 'z-score':
                cell_format = red_bg_format if value > 0 else data_format
            else:
                cell_format = data_format
            worksheet.write(row + 1, col, value, cell_format)

    # Write additional table for threshold value, mean, and standard deviation
    worksheet.write(len(booking_counts_complete) + 2, 0, 'Threshold', heading_format)
    worksheet.write(len(booking_counts_complete) + 2, 1, threshold_value, data_format)

    worksheet.write(len(booking_counts_complete) + 3, 0, 'Mean', heading_format)
    worksheet.write(len(booking_counts_complete) + 3, 1, mean_bookings, data_format)

    worksheet.write(len(booking_counts_complete) + 4, 0, 'Standard Deviation', heading_format)
    worksheet.write(len(booking_counts_complete) + 4, 1, std_bookings, data_format)

    workbook.close()
    print(f"Excel file saved successfully: {excel_filepath}")

# Example file path, replace this with your actual file path when you use the code
file_path = 'C:\\Users\\Jyothesh karnam\\Desktop\\collaborative application development\\Data Files\\D19.csv'
plot_combined_booking_activity_analysis_and_display_table(file_path)

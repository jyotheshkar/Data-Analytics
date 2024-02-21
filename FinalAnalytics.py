import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import zscore
import numpy as np
import os
import xlsxwriter

def plot_combined_booking_activity_analysis_and_display_table(file_path):
    # Reading the CSV file, skipping the first row if necessary
    df = pd.read_csv(file_path, skiprows=1)
    
    # Converting 'Created Date' to datetime format
    df['Created Date'] = pd.to_datetime(df['Created Date'], format='%d/%m/%Y')
    
    # Finding the first and last dates
    start_date = df['Created Date'].min()
    end_date = df['Created Date'].max()
    
    # Calculating the nearest Monday
    days_to_monday = (start_date.weekday() - 0) % 7
    first_monday = start_date - pd.Timedelta(days=days_to_monday)
    
    # Calculating sequential week numbers
    df['Sequential Week Number'] = ((df['Created Date'] - first_monday).dt.days // 7) + 1
    all_weeks = pd.DataFrame({'Sequential Week Number': range(1, ((end_date - first_monday).days // 7) + 2)})
    
    # Grouping data by week
    booking_counts = df.groupby('Sequential Week Number').size().reset_index(name='Number of Bookings')
    booking_counts_complete = pd.merge(all_weeks, booking_counts, on='Sequential Week Number', how='left').fillna(0)
    
    # Calculating z-scores for booking counts
    booking_counts_complete['z-score'] = zscore(booking_counts_complete['Number of Bookings']).round(4)
    
    # Plotting
    plt.style.use('classic')
    fig, ax = plt.subplots(figsize=(16, 12))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    ax.bar(booking_counts_complete['Sequential Week Number'], booking_counts_complete['Number of Bookings'], color='skyblue', alpha=0.7, width=1)
    
    # Calculating mean, std, and threshold
    mean_bookings = booking_counts_complete['Number of Bookings'].mean()
    std_bookings = booking_counts_complete['Number of Bookings'].std()
    threshold_value = mean_bookings + std_bookings
    offset = max(booking_counts_complete['Number of Bookings']) * 0.02
    
    # Annotating bars
    for i, row in booking_counts_complete.iterrows():
        color = 'red' if row['Number of Bookings'] > threshold_value else 'black'
        ax.text(row['Sequential Week Number'], row['Number of Bookings'] + offset, f"{int(row['Number of Bookings'])}", ha='center', va='center', color=color, fontsize=8)
    
    ax.set_xlabel('Week No.')
    ax.set_ylabel('Number of Bookings')
    filename = os.path.basename(file_path)
    ax.set_title(f'Booking Count by Week ({filename})')
    
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_xlim(0.5, len(all_weeks) + 0.5)
    ax.set_ylim(0, max(booking_counts_complete['Number of Bookings']) + (5 * offset))
    
    ax.axhline(y=threshold_value, color='darkorange', linestyle='--', linewidth=2)
    ax.text(len(all_weeks) + 0.5, threshold_value, f"Threshold: {threshold_value:.2f}", va='center', ha='left', color='orange')
    
    plt.tight_layout()
    
    # Dynamic image saving
    base_directory = os.path.dirname(file_path)
    filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
    image_directory = os.path.join(base_directory, "Images")
    os.makedirs(image_directory, exist_ok=True)
    image_file_path = os.path.join(image_directory, f"{filename_without_extension}.png")
    plt.savefig(image_file_path, dpi=300, bbox_inches='tight')
    print(f"Plot image saved at: {image_file_path}")
    
    plt.show()  # Displays the plot after saving
    
    # Excel report generation
    excel_filename = f"datareport_{filename_without_extension}.xlsx"
    excel_folder = os.path.join(base_directory, "reports")
    os.makedirs(excel_folder, exist_ok=True)
    excel_filepath = os.path.join(excel_folder, excel_filename)
    workbook = xlsxwriter.Workbook(excel_filepath)
    worksheet = workbook.add_worksheet()

    heading_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black', 'border': 1})
    data_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000'})

    headings = ['Sequential Week Number', 'Number of Bookings', 'z-score']
    worksheet.write_row('A1', headings, heading_format)
    for row, data in enumerate(booking_counts_complete.values, start=1):
        worksheet.write_row(f'A{row+1}', data, data_format)

    workbook.close()
    print(f"Excel file saved successfully: {excel_filepath}")

# Specify your CSV file path
file_path = 'C:\\Users\\Jyothesh karnam\\Desktop\\collaborative application development\\Data Files\\SRM23.csv'
plot_combined_booking_activity_analysis_and_display_table(file_path)

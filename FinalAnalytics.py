import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import zscore
import os
import xlsxwriter

def plot_combined_booking_activity_analysis_and_display_table(file_path):
    df = pd.read_csv(file_path, skiprows=1)
    df['Created Date'] = pd.to_datetime(df['Created Date'], format='%d/%m/%Y')
    start_date = df['Created Date'].min()
    end_date = df['Created Date'].max()
    days_to_monday = (start_date.weekday() - 0) % 7
    first_monday = start_date - pd.Timedelta(days=days_to_monday)
    df['Sequential Week Number'] = ((df['Created Date'] - first_monday).dt.days // 7) + 1
    all_weeks = pd.DataFrame({'Sequential Week Number': range(1, ((end_date - first_monday).days // 7) + 2)})
    booking_counts = df.groupby('Sequential Week Number').size().reset_index(name='Number of Bookings')
    booking_counts_complete = pd.merge(all_weeks, booking_counts, on='Sequential Week Number', how='left').fillna(0)
    
    # Ensuring standard deviation includes weeks with zero bookings correctly
    mean_bookings = booking_counts_complete['Number of Bookings'].mean()
    # Manual calculation of standard deviation to ensure clarity
    diff_squared = (booking_counts_complete['Number of Bookings'] - mean_bookings) ** 2
    variance = diff_squared.mean()
    std_bookings = variance ** 0.5
    threshold_value = mean_bookings + std_bookings

    booking_counts_complete['z-score'] = ((booking_counts_complete['Number of Bookings'] - mean_bookings) / std_bookings).round(4)
    
    plt.style.use('classic')
    fig, ax = plt.subplots(figsize=(16, 12))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')
    bars = ax.bar(booking_counts_complete['Sequential Week Number'], booking_counts_complete['Number of Bookings'], color='skyblue', alpha=0.7, width=1)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{int(height)}',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha='center', va='bottom')
    for i, row in booking_counts_complete.iterrows():
        color = 'red' if row['Number of Bookings'] > threshold_value else 'black'
        ax.scatter(row['Sequential Week Number'], row['Number of Bookings'], color=color)
    ax.set_xlabel('Week No.')
    ax.set_ylabel('Number of Bookings')
    filename = os.path.basename(file_path)
    ax.set_title(f'Booking Count by Week ({filename})')
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    total_weeks = len(all_weeks)
    if total_weeks > 50:
        ax.set_xticks(range(1, total_weeks + 1, 5))
    else:
        ax.set_xticks(range(1, total_weeks + 1))
    ax.set_xlim(0.5, total_weeks + 0.5)
    ax.axhline(y=threshold_value, color='darkorange', linestyle='--', linewidth=2)
    ax.text(len(all_weeks) + 0.5, threshold_value, f"Threshold: {threshold_value:.2f}", va='center', ha='left', color='white', backgroundcolor='black', bbox=dict(facecolor='black', edgecolor='orange', boxstyle="round,pad=0.3"))
    plt.tight_layout()
    base_directory = os.path.dirname(file_path)
    filename_without_extension = os.path.splitext(os.path.basename(file_path))[0]
    image_directory = os.path.join(base_directory, "Images")
    os.makedirs(image_directory, exist_ok=True)
    image_file_path = os.path.join(image_directory, f"{filename_without_extension}.png")
    plt.savefig(image_file_path, dpi=300, bbox_inches='tight')
    print(f"Plot image saved at: {image_file_path}")
    plt.show()
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
        for col, value in enumerate(data):
            format_to_use = data_format
            if col == 2 and value > 1:
                format_to_use = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000', 'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            worksheet.write(row, col, value, format_to_use)
    summary_row_start = len(booking_counts_complete) + 3
    summary_data = [("Threshold", threshold_value), ("Mean", mean_bookings), ("Standard Deviation", std_bookings)]
    summary_heading_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': 'black', 'border': 1})
    summary_data_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0.0000'})
    for i, (label, value) in enumerate(summary_data):
        worksheet.write(summary_row_start + i, 0, label, summary_heading_format)
        worksheet.write(summary_row_start + i, 1, value, summary_data_format)
    workbook.close()
    print(f"Excel file saved successfully: {excel_filepath}")

# Specify your CSV file path here
file_path = ''
plot_combined_booking_activity_analysis_and_display_table(file_path)

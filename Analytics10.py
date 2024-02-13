import pandas as pd
import matplotlib.pyplot as plt
from scipy.stats import zscore

def plot_combined_booking_activity_analysis(file_path):
    # Read data from CSV
    df = pd.read_csv(file_path)
    
    # Strip leading/trailing spaces from column names
    df.columns = df.columns.str.strip()
    
    if 'Created Date' not in df.columns:
        print("Column 'Created Date' not found. Please check the column names.")
        return
    
    # Convert 'Created Date' to datetime format, specifying dayfirst due to DD/MM/YYYY format
    df['Created Date'] = pd.to_datetime(df['Created Date'], dayfirst=True)
    
    # Sort DataFrame by 'Created Date'
    df.sort_values(by='Created Date', inplace=True)
    
    # Extract sequential week number from 'Created Date'
    df['Sequential Week Number'] = df['Created Date'].dt.isocalendar().week
    
    # Generate a complete sequence of weeks from the minimum to the maximum
    complete_weeks = range(df['Sequential Week Number'].min(), df['Sequential Week Number'].max() + 1)
    
    # Initialize a DataFrame to represent all weeks
    all_weeks_df = pd.DataFrame(complete_weeks, columns=['Sequential Week Number'])
    
    # Count bookings per week
    booking_counts = df.groupby('Sequential Week Number')['Created Date'].count().reset_index(name='Number of Bookings')
    
    # Merge to ensure all weeks are represented, filling missing weeks with 0 bookings
    booking_counts = all_weeks_df.merge(booking_counts, on='Sequential Week Number', how='left').fillna(0)
    
    # Calculate Z-scores of booking counts
    booking_counts['z-score'] = zscore(booking_counts['Number of Bookings'].astype(float), ddof=0)
    
    # Map 'Sequential Week Number' to a new sequence starting from 1
    booking_counts['Week Sequence'] = range(1, len(booking_counts) + 1)

    plt.style.use('dark_background')
    fig, ax = plt.subplots(figsize=(16, 12))

    # Use 'Week Sequence' for plotting on X-axis
    ax.bar(booking_counts['Week Sequence'], booking_counts['Number of Bookings'], color='grey', alpha=0.5, width=1)

    # Calculations include weeks with 0 bookings
    mean_bookings = booking_counts['Number of Bookings'].mean()
    std_bookings = booking_counts['Number of Bookings'].std(ddof=0)  # Ensure this line correctly calculates including zeros
    threshold_value = mean_bookings + std_bookings

    # Calculate offset for text annotation based on range of 'Number of Bookings'
    offset = (max(booking_counts['Number of Bookings']) - min(booking_counts['Number of Bookings'])) * 0.02

    # Annotate each bar with booking count and z-score
    for i, row in booking_counts.iterrows():
        color = 'green' if row['Number of Bookings'] > threshold_value else 'red'
        ax.scatter(row['Week Sequence'], row['Number of Bookings'], color=color, zorder=5)
        ax.text(row['Week Sequence'], row['Number of Bookings'] + offset, f"{row['Number of Bookings']}", 
                ha='center', va='bottom', color='white', fontsize=8, 
                bbox=dict(facecolor='grey', edgecolor='none', boxstyle="round,pad=0.1"))
        ax.text(row['Week Sequence'], row['Number of Bookings'] + 2 * offset, f"Z={row['z-score']:.4f}", rotation=90,
                ha='center', va='bottom', color='white', fontsize=8, 
                bbox=dict(facecolor='grey', edgecolor='none', boxstyle="round,pad=0.1"))

    # Draw threshold line and update legend
    ax.axhline(y=threshold_value, color='orange', linestyle='--', label=f'Threshold: {threshold_value:.4f}')

    legend_text = f'Threshold: {threshold_value:.4f}\nMean: {mean_bookings:.4f}\nSD: {std_bookings:.4f}'
    ax.legend([legend_text], loc='upper left')

    # Configure axis labels and title
    ax.set_xlabel('Sequential Week Number')
    ax.set_ylabel('Number of Bookings')
    ax.set_title('Combined Booking Activity Analysis')
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_xticks(booking_counts['Week Sequence'])
    ax.set_xticklabels(booking_counts['Week Sequence'], rotation=45)
    ax.set_ylim(0, max(booking_counts['Number of Bookings']) + (5 * offset))
    
    plt.tight_layout()
    plt.show()


# Specify the correct path to your CSV file
file_path = ""
plot_combined_booking_activity_analysis(file_path)
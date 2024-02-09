
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.ticker as ticker

def calculate_threshold_mean_std(booking_counts, num_std=1):
    """Calculate a threshold as mean plus a number of standard deviations."""
    mean_bookings = booking_counts['Number of Bookings'].mean()
    std_bookings = booking_counts['Number of Bookings'].std()
    threshold = mean_bookings + num_std * std_bookings
    return threshold

def plot_booking_activity(file_path):
    """Plot the booking activity from a CSV file, setting the threshold using mean and standard deviation."""
    # Read and preprocess data
    df = pd.read_csv(file_path, skiprows=1)
    df['Created Date'] = pd.to_datetime(df.iloc[:, 1], format='%d/%m/%Y')
    start_date = df['Created Date'].min()
    days_to_monday = (start_date.weekday() - 0) % 7
    first_monday = start_date - pd.Timedelta(days=days_to_monday)
    df['Sequential Week Number'] = ((df['Created Date'] - first_monday).dt.days // 7) + 1
    
    booking_counts = df.groupby('Sequential Week Number').size().reset_index(name='Number of Bookings')
    threshold = calculate_threshold_mean_std(booking_counts)

    plt.style.use('dark_background')
    fig, ax = plt.subplots(figsize=(16, 10))

    # Plot the booking counts
    ax.bar(booking_counts['Sequential Week Number'], booking_counts['Number of Bookings'], color='grey', alpha=0.5, width=1, label='Booking Counts')
    
    # Dummy scatter points for legend
    ax.scatter([], [], color='green', label='Above Threshold')
    ax.scatter([], [], color='red', label='Below Threshold')

    # Plot each data point with color based on threshold comparison
    for i, row in booking_counts.iterrows():
        color = 'red' if row['Number of Bookings'] < threshold else 'green'
        ax.scatter(row['Sequential Week Number'], row['Number of Bookings'], color=color, zorder=5)
        # Annotate the exact number of bookings
        ax.annotate(row['Number of Bookings'], (row['Sequential Week Number'], row['Number of Bookings']),
                    textcoords="offset points", xytext=(0,10), ha='center', color='black', fontsize=8,
                    bbox=dict(facecolor='white', edgecolor='none', boxstyle="round,pad=0.3"))

    # Draw and annotate the threshold line horizontally
    ax.axhline(y=threshold, color='orange', linestyle='--')
    ax.annotate(f'Threshold: {threshold:.2f}', xy=(1, threshold), xycoords=('axes fraction', 'data'),
                textcoords="offset points", xytext=(-10,10), ha='right', color='orange', fontsize=12,
                bbox=dict(facecolor='black', alpha=0.5))

    ax.set_xlabel('Sequential Week Number')
    ax.set_ylabel('Number of Bookings')
    ax.set_title('Booking Activity and Marketing Impact Analysis')
    ax.legend(loc='upper right')

    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)
    ax.set_xticks(range(1, booking_counts['Sequential Week Number'].max() + 1))
    ax.set_xticklabels(range(0, booking_counts['Sequential Week Number'].max()))

    plt.tight_layout()
    plt.show()

# Example file path
file_path = ""
plot_booking_activity(file_path)



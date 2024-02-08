import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.ticker as ticker

def determine_threshold(booking_counts, percentile=80):
    return np.percentile(booking_counts['Number of Bookings'], percentile)

def plot_booking_activity(file_path):
    # Step 1: Read the data
    df = pd.read_csv(file_path, skiprows=1)
    df['Created Date'] = pd.to_datetime(df.iloc[:, 1], format='%d/%m/%Y')
    num_weeks = ((df['Created Date'] - df['Created Date'].min()) / pd.Timedelta(weeks=1)).astype(int) + 1
    df['Time Interval'] = num_weeks
    booking_counts = df.groupby('Time Interval').size().reset_index(name='Number of Bookings')
    threshold = determine_threshold(booking_counts)

    # Calculate percentage change for booking counts and limit to 100%
    booking_counts['Percentage Change'] = booking_counts['Number of Bookings'].pct_change().apply(lambda x: min(x, 1.0)) * 100

    plt.figure(figsize=(12, 8))

    # Plot peaks above and below threshold
    peaks_above_threshold = booking_counts[booking_counts['Number of Bookings'] > threshold]
    peaks_below_threshold = booking_counts[booking_counts['Number of Bookings'] <= threshold]
    plt.scatter(peaks_above_threshold['Time Interval'], peaks_above_threshold['Number of Bookings'], color='red', label='Peaks > Threshold', marker='o')
    plt.scatter(peaks_below_threshold['Time Interval'], peaks_below_threshold['Number of Bookings'], color='black', label='Peaks <= Threshold', marker='o')

    plt.plot(booking_counts['Time Interval'], booking_counts['Number of Bookings'], linestyle='-', color='black')

    # Annotate the number of bookings at each peak with black background
    for i, txt in enumerate(booking_counts['Number of Bookings']):
        plt.annotate(txt, (booking_counts['Time Interval'][i], booking_counts['Number of Bookings'][i]), 
                     textcoords="offset points", xytext=(0,10), ha='center', 
                     fontsize=8, color='white', bbox=dict(facecolor='black', edgecolor='none', boxstyle="round"))

    # Annotate percentage changes on slopes, color-coded and vertical
    for i in range(1, len(booking_counts)):
        pct_change = booking_counts.iloc[i]['Percentage Change']
        if not np.isnan(pct_change):
            mid_point_x = (booking_counts.iloc[i]['Time Interval'] + booking_counts.iloc[i-1]['Time Interval']) / 2
            mid_point_y = (booking_counts.iloc[i]['Number of Bookings'] + booking_counts.iloc[i-1]['Number of Bookings']) / 2
            color = 'green' if pct_change > 0 else 'red'
            plt.text(mid_point_x, mid_point_y, f"{pct_change:.2f}%", rotation=90, 
                     ha='center', va='bottom' if pct_change > 0 else 'top',
                     fontsize=8, color=color)

    plt.axhline(y=threshold, color='orange', linestyle='--', label=f'Threshold: {threshold}')

    # Set x-axis to show date ranges with vertical labels (Adjust as needed based on your dataset)
    plt.xticks(booking_counts['Time Interval'], rotation='vertical', ha="center")

    plt.gca().yaxis.set_major_locator(ticker.MaxNLocator(integer=True))
    plt.legend()
    plt.title('Booking Activity Over Time')
    plt.xlabel('Time Interval (Weeks from Start)')
    plt.ylabel('Number of Bookings')
    plt.grid(True)
    plt.tight_layout()
    plt.show()

# Replace the file_path with the absolute path to your CSV file
file_path = ""
plot_booking_activity(file_path)
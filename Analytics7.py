import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.ticker as ticker

def determine_threshold(booking_counts, percentile=80):
    return np.percentile(booking_counts['Number of Bookings'], percentile)

def plot_booking_activity(file_path):
    df = pd.read_csv(file_path, skiprows=1)
    df['Created Date'] = pd.to_datetime(df.iloc[:, 1], format='%d/%m/%Y')
    num_weeks = ((df['Created Date'] - df['Created Date'].min()) / np.timedelta64(1, 'W')).astype(int)
    df['Time Interval'] = num_weeks
    booking_counts = df.groupby('Time Interval').size().reset_index(name='Number of Bookings')

    threshold = determine_threshold(booking_counts)

    plt.style.use('dark_background')
    fig, ax = plt.subplots(figsize=(16, 10))

    # Bar chart for booking counts
    ax.bar(booking_counts['Time Interval'], booking_counts['Number of Bookings'], color='grey', alpha=0.5, label='Booking Counts')

    # Scatter plot for individual data points
    ax.scatter(booking_counts['Time Interval'], booking_counts['Number of Bookings'], color='white', zorder=5)

    for i in range(1, len(booking_counts)):
        start = booking_counts.iloc[i - 1]
        end = booking_counts.iloc[i]
        pct_change = ((end['Number of Bookings'] - start['Number of Bookings']) / booking_counts['Number of Bookings'].sum()) * 100
        color = 'green' if pct_change > 0 else 'red'
        ax.plot([start['Time Interval'], end['Time Interval']], [start['Number of Bookings'], end['Number of Bookings']], color=color, lw=2)
        
        # Determine annotation rotation based on the slope of the line
        rotation = 'vertical' if start['Number of Bookings'] != end['Number of Bookings'] else 'horizontal'
        mid_point = (start['Time Interval'] + end['Time Interval']) / 2
        ax.annotate(f"{pct_change:.2f}%", (mid_point, (start['Number of Bookings'] + end['Number of Bookings']) / 2), color='white', 
                    textcoords="offset points", xytext=(0,10), ha='center', fontsize=8, 
                    bbox=dict(facecolor=color, alpha=0.5), rotation=rotation)

    # Annotate each data point with the number of bookings
    for i, row in booking_counts.iterrows():
        ax.annotate(row['Number of Bookings'], (row['Time Interval'], row['Number of Bookings']), 
                    textcoords="offset points", xytext=(0,10), ha='center', color='black', fontsize=8, 
                    bbox=dict(facecolor='white', edgecolor='none', boxstyle="round,pad=0.3"))

    ax.axhline(y=threshold, color='orange', linestyle='--', label='Threshold')

    ax.set_xlabel('Time Interval (Weeks from Start)')
    ax.set_ylabel('Number of Bookings')
    ax.set_title('Booking Activity and Marketing Impact Analysis')
    ax.legend()
    ax.grid(True, color='gray', linestyle='--', linewidth=0.5)

    # Ensure all numbers are shown on the x-axis by setting ticks at every unique time interval value
    all_time_intervals = np.arange(booking_counts['Time Interval'].min(), booking_counts['Time Interval'].max() + 1)
    ax.set_xticks(all_time_intervals)
    ax.set_xticklabels(all_time_intervals, rotation=90)

    plt.tight_layout()
    plt.show()

# Your file path
file_path = " give your file path here "
plot_booking_activity(file_path)

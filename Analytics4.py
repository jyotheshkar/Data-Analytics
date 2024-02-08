import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import matplotlib.ticker as ticker

# Step 1: Read the data
# Note: You will need to update the 'file_path' to the correct path where your file is located.
file_path = ""
df = pd.read_csv(file_path, skiprows=1)

# Assuming 'Created Date' is in the second column
df['Created Date'] = pd.to_datetime(df.iloc[:, 1], format='%d/%m/%Y')

# Calculate the number of weeks between the minimum date and each 'Created Date'
num_weeks = ((df['Created Date'] - df['Created Date'].min()) / pd.Timedelta(weeks=1)).astype(int) + 1

# Create a new column for the time interval based on the calculated weeks
df['Time Interval'] = num_weeks

# Group by time interval and count the number of bookings at each interval
booking_counts = df.groupby('Time Interval').size().reset_index(name='Number of Bookings')

# Generate date ranges for each week
date_range_labels = []
for week in booking_counts['Time Interval']:
    start_date = df['Created Date'].min() + pd.Timedelta(weeks=week-1)
    end_date = start_date + pd.Timedelta(days=6)
    date_range_labels.append(f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}")

# Plotting
plt.figure(figsize=(10, 6))

# Plot linear curve in green colour
fit = np.polyfit(booking_counts['Time Interval'], booking_counts['Number of Bookings'], 1)
linear_curve = np.poly1d(fit)
plt.plot(booking_counts['Time Interval'], linear_curve(booking_counts['Time Interval']), linestyle='--', color='green', label='Linear Curve')

# Plot peaks above and below threshold
threshold = 20
peaks_above_threshold = booking_counts[booking_counts['Number of Bookings'] > threshold]
peaks_below_threshold = booking_counts[booking_counts['Number of Bookings'] <= threshold]
plt.scatter(peaks_above_threshold['Time Interval'], peaks_above_threshold['Number of Bookings'], color='red', label='Peaks > Threshold', marker='o')
plt.scatter(peaks_below_threshold['Time Interval'], peaks_below_threshold['Number of Bookings'], color='black', label='Peaks <= Threshold', marker='o')

# Plot line graph
plt.plot(booking_counts['Time Interval'], booking_counts['Number of Bookings'], linestyle='-', color='black')

# Annotate each data point with the number of bookings
for i, txt in enumerate(booking_counts['Number of Bookings']):
    plt.annotate(txt, (booking_counts['Time Interval'][i], booking_counts['Number of Bookings'][i]), textcoords="offset points", xytext=(0,10), ha='center', 
                 fontsize=8, color='white', bbox=dict(facecolor='black', edgecolor='none', boxstyle="round"))

# Add a horizontal line at the threshold level
plt.axhline(y=threshold, color='orange', linestyle='--', label='Threshold')

# Set x-axis to show date ranges
plt.xticks(booking_counts['Time Interval'], date_range_labels, rotation=45, ha="right")

# Ensure all booking numbers are visible on y-axis
plt.gca().yaxis.set_major_locator(ticker.MaxNLocator(integer=True))

# Show legend
plt.legend()

# Set labels and title
plt.title('Booking Activity Over Time')
plt.xlabel('Time Interval (Date Range)')
plt.ylabel('Number of Bookings')
plt.grid(True)

# Adjust layout and display the plot
plt.tight_layout()
plt.show()

# Print column names to identify the correct column for advertisement
print(df.columns)
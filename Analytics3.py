import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# Step 1: Read the data
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

# Plot line graph
plt.figure(figsize=(10, 6))

# Plot linear curve in green colour
fit = np.polyfit(booking_counts['Time Interval'], booking_counts['Number of Bookings'], 1)
linear_curve = np.poly1d(fit)
plt.plot(booking_counts['Time Interval'], linear_curve(booking_counts['Time Interval']), linestyle='--', color='green', label='Linear Curve')

# Identify peaks exceeding a certain threshold
threshold = 30  # Set your desired threshold
peaks_above_threshold = booking_counts[booking_counts['Number of Bookings'] > threshold]
peaks_below_threshold = booking_counts[booking_counts['Number of Bookings'] <= threshold]

# Plot peaks above threshold in red
plt.scatter(peaks_above_threshold['Time Interval'], peaks_above_threshold['Number of Bookings'], color='red', label='Peaks > Threshold', marker='o')

# Plot peaks below threshold in black
plt.scatter(peaks_below_threshold['Time Interval'], peaks_below_threshold['Number of Bookings'], color='black', label='Peaks <= Threshold', marker='o')

# Plot line graph
plt.plot(booking_counts['Time Interval'], booking_counts['Number of Bookings'], linestyle='-', color='black')

# Add a horizontal line at the threshold level
plt.axhline(y=threshold, color='orange', linestyle='--', label='Threshold')

# Show legend
plt.legend()

# Set labels and title
plt.title('Booking Activity Over Time')
plt.xlabel('Time Interval (weeks)')
plt.ylabel('Number of Bookings')
plt.grid(True)

# Display the plot
plt.show()

# Print column names to identify the correct column for advertisement
print(df.columns)
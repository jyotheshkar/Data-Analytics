import pandas as pd

# Update the file path to your specific location
file_path = ""

# Load the CSV data into a Pandas DataFrame, skip the first row, and infer the column names
df = pd.read_csv(file_path, skiprows=1, header=None)

# Assign meaningful column names based on your data
df.columns = ['BookingReference', 'Created Date', 'Reference', 'Attendee Status', 'Attended']

# Basic statistics summary
print("\nSummary Statistics:")
print(df.describe())

# Count of unique values in 'Attendee Status'
attendee_status_counts = df['Attendee Status'].value_counts()
print("\nAttendee Status Counts:")
print(attendee_status_counts)

# Count of unique values in 'Attended'
attended_counts = df['Attended'].value_counts()

# Count of missing values in 'Attended'
missing_attendance_counts = df['Attended'].isna().sum()
print("\nAttended Counts:")
print(attended_counts)
print(f"Missing/Empty: There are {missing_attendance_counts} entries with missing or empty values in the Attended column.")

# Group by 'Attendee Status' and calculate average attendance as a percentage
avg_attendance_by_status = df.groupby('Attendee Status')['Attended'].apply(lambda x: (x == 'Yes').mean() * 100)
print("\nAverage Attendance by Attendee Status (in Percentage):")
print(avg_attendance_by_status)
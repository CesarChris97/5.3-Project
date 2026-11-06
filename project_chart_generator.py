import csv
from datetime import datetime
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference

# Student ID definition and file paths
student_id = "CESORT4471"
file_path = "ZooData.csv"
excel_path = "final.xlsx"

# Function: convertData
# Description: Converts temperature from Fahrenheit to Celsius
def convertData(val):
    return (val - 32) * 5 / 9

# Function: insertData
# Description: Appends a comma-separated data string to the CSV file using try-except error handling
def insertData(path, data_string):
    try:
        with open(path, mode="a", newline="", encoding="utf-8") as file:
            file.write(data_string + "\n")
    except Exception as e:
        print(f"Error writing to file: {e}")

# Function: viewData
# Description: Reads and displays the contents of the CSV file using minimal read permissions and try-except error handling
def viewData(path):
    try:
        print(f"The file {path}")
        with open(path, mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            for row in reader:
                if row:
                    print(",".join(row))
    except FileNotFoundError:
        print(f"Error: The file '{path}' was not found.")
    except Exception as e:
        print(f"Error reading file: {e}")

# Function: getInput
# Description: Prompts user for inputs, converts data, saves to CSV via insertData, and displays confirmation
def getInput():
    try:
        count = int(input("How many entries are you inputting?\n"))
        for _ in range(count):
            date_str = input("\nEnter a date:\n")
            temp_val = float(input("Enter the highest temp for the inputted date:\n"))
            
            # Convert temperature
            converted_val = convertData(temp_val)
            
            # Format row string: date,original_temp,converted_temp
            formatted_data = f"{date_str},{temp_val},{converted_val}"
            
            # Append data to ZooData.csv using insertData
            insertData(file_path, formatted_data)
            
            # Confirm save with timestamp
            print(f"The following was saved at {datetime.now()} :")
            print(formatted_data)
    except ValueError:
        print("Error: Please enter valid numerical values for entry count and temperature.")

# Function: createChart
# Description: Reads CSV data, writes to final.xlsx using openpyxl, generates line/bar chart with axis labels and title, and saves the file.
# Arguments: path (str) - path to CSV data file, chart_type (str) - 'bar' or 'line'
# Return Value: None
def createChart(path, chart_type):
    try:
        print("\nChoose data source for chart:")
        print("1 Fahrenheit (Initial Data)")
        print("2 Celsius (Converted Data)")
        data_choice = input().strip()

        if data_choice == "1":
            value_col_idx = 1  # 0-indexed column 1 (Fahrenheit)
            y_label = "Temperature (°F)"
        elif data_choice == "2":
            value_col_idx = 2  # 0-indexed column 2 (Celsius)
            y_label = "Temperature (°C)"
        else:
            print("Invalid data choice defaulting to Fahrenheit.")
            value_col_idx = 1
            y_label = "Temperature (°F)"

        # Read CSV file contents
        dates = []
        values = []
        with open(path, mode="r", encoding="utf-8") as file:
            reader = csv.reader(file)
            for row in reader:
                if row:
                    dates.append(row[0])
                    values.append(float(row[value_col_idx]))

        # Create new Excel workbook and sheet
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Report"

        # Write Headers
        ws.append(["Date", y_label])

        # Write Data Rows
        for d, v in zip(dates, values):
            ws.append([d, v])

        # Configure Chart Type
        if chart_type.lower() == "bar":
            chart = BarChart()
        else:
            chart = LineChart()

        # Set Title as "<student ID> <current date>"
        current_date_str = datetime.now().strftime("%m/%d/%Y")
        chart.title = f"{student_id} {current_date_str}"
        chart.x_axis.title = "Date"
        chart.y_axis.title = y_label

        # Set Chart Data and Categories References
        data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(values) + 1)
        cats_ref = Reference(ws, min_col=1, min_row=2, max_row=len(dates) + 1)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        # Place Chart in Worksheet and Save Workbook
        ws.add_chart(chart, "D2")
        wb.save(excel_path)
        print(f"Successfully generated report and chart in {excel_path}!")

    except FileNotFoundError:
        print(f"Error: The file '{path}' was not found.")
    except Exception as e:
        print(f"Error creating chart: {e}")

# Function: generateReport
# Description: Prompts user to select chart type ('bar' or 'line') and invokes createChart.
# Arguments: path (str) - path to CSV data file
# Return Value: None
def generateReport(path):
    print("\nChoose graph type to create:")
    print("1 Bar Chart")
    print("2 Line Chart")
    graph_choice = input().strip()

    if graph_choice == "1":
        createChart(path, "bar")
    elif graph_choice == "2":
        createChart(path, "line")
    else:
        print("Error: Invalid graph choice selected.")

# Main Program Execution & Menu Flow
print(f"{student_id}'s Spreadsheet Automation Menu")

menu_options = {
    "1": "Input Data",
    "2": "View Current Data",
    "3": "Generate Report"
}

print("Choose a number from the following options")
for key, value in menu_options.items():
    print(f"{key} {value}")

choice = input().strip()

# Menu Logic & Routing
if choice == "1":
    print(f"You selected 1 at {datetime.now()}")
    getInput()
elif choice == "2":
    print(f"You selected 2 at {datetime.now()}")
    viewData(file_path)
elif choice == "3":
    print(f"You selected 3 at {datetime.now()}")
    generateReport(file_path)
else:
    print("Error: Invalid choice selected")

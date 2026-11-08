from datetime import datetime
from pathlib import Path
import csv
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference

CSV_FILE = Path(__file__).with_name('ZooData.csv')


# Converts a Fahrenheit temperature to Celsius.
def convertData(temperature):
    return (temperature - 32) * 5 / 9


# Appends a comma-separated record to the CSV file.
def insertData(file_path, data):
    try:
        with open(file_path, 'a', encoding='utf-8', newline='') as file:
            file.write(data + '\n')
        return True
    except OSError as error:
        print(f'Error saving data: {error}')
        return False


# Displays the CSV file location and all of its saved records.
def viewData(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            print(f'The file {Path(file_path).resolve()} contains:')
            contents = file.read()
            print(contents if contents else 'No entries have been saved yet.', end='' if contents else '\n')
    except OSError as error:
        print(f'Error reading data: {error}')


# Collects dates and temperatures, converts them, and saves each entry.
def getInput():
    try:
        number_of_entries = int(input('How many entries would you like to add? '))
        if number_of_entries <= 0:
            print('Please enter a number greater than zero.')
            return

        for entry in range(number_of_entries):
            date = input('Enter a date (MM/DD/YYYY): ').strip()
            temperature = float(input('Enter the highest temp for the inputted date: '))
            converted = convertData(temperature)
            data = f'{date},{temperature:g},{converted}'

            if insertData(CSV_FILE, data):
                print(f'The following data was saved at {datetime.now()}')
                print(data)
    except ValueError:
        print('Error: Please enter a valid number for the count or temperature.')
    except (EOFError, KeyboardInterrupt):
        print('\nData entry was canceled.')

# Creates an Excel chart using temperature data from a CSV file.
# Arguments: file_path (CSV location), chart_type ("bar" or "line").
# Returns: None. Saves the chart and data to final.xlsx.
def createChart(file_path, chart_type):
    source = input(
        "Choose your data (1 = Fahrenheit, 2 = Celsius): "
    ).strip()

    if source not in ('1', '2'):
        print("Invalid data selection.")
        return

    try:
        with open(file_path, 'r', encoding='utf-8-sig', newline='') as file:
            reader = csv.reader(file)
            records = list(reader)

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Temperature Report"

        unit = "Fahrenheit" if source == '1' else "Celsius"

        sheet.append(["Date", unit])

        for record in records:
            if len(record) < 3:
                continue

            try:
                date = record[0].strip()
                temperature = float(record[int(source)].strip())
                sheet.append([date, temperature])
            except ValueError:
                continue

        if sheet.max_row < 2:
            print("No valid temperature data found.")
            return

        if chart_type == 'bar':
            chart = BarChart()
        else:
            chart = LineChart()

        chart.title = f"coryou8501 {datetime.now():%m/%d/%Y}"
        chart.x_axis.title = "Date"
        chart.y_axis.title = f"Temperature ({unit})"

        
        values = Reference(
            sheet,
            min_col=2,
            min_row=1,
            max_row=sheet.max_row
        )

        labels = Reference(
            sheet,
            min_col=1,
            min_row=2,
            max_row=sheet.max_row
        )

        chart.add_data(values, titles_from_data=True)
        chart.set_categories(labels)

        chart.x_axis.title = "Date"
        chart.y_axis.title = f"Temperature ({unit})"

        chart.width = 18
        chart.height = 10
        chart.legend = None

        sheet.add_chart(chart, "D2")


        output_file = Path(file_path).with_name("final.xlsx")
        workbook.save(output_file)

        print(f"Report created successfully: {output_file}")

    except OSError as error:
        print(f"Error creating report: {error}")

# Asks the user which chart to create and generates the report.
# Argument: file_path (location of the CSV data file).
# Returns: None. Calls createChart() with the selected chart type.
def generateReport(file_path):
    print("\nGenerate Temperature Report")
    print("1 Bar Chart")
    print("2 Line Chart")

    choice = input("Choose your chart type (1 or 2): ").strip()

    if choice == '1':
        createChart(file_path, 'bar')
    elif choice == '2':
        createChart(file_path, 'line')
    else:
        print("Invalid chart selection. Please choose 1 or 2.")


# Displays the menu and runs the function selected by the user.
def main():
    while True:
        print('\n' + '=' * 52)
        print("coryou8501's Spreadsheet Automation Menu")
        print('Choose a number from the following options')
        print('1 Input Data')
        print('2 View Current Data')
        print('3 Generate Report')
        print('4 Exit')

        choice = input('Enter your choice: ').strip()
        if choice in ('1', '2', '3', '4'):
            print(f'You selected {choice} at {datetime.now()}')
        if choice == '1':
            getInput()
        elif choice == '2':
            viewData(CSV_FILE)
        elif choice == '3':
            generateReport(CSV_FILE)
        elif choice == '4':
            print('Goodbye!')
            break
        else:
            print('Error: Invalid choice selected.')


# Starts the spreadsheet automation program.
if __name__ == '__main__':
    main()

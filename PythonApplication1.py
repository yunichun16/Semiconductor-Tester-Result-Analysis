#this is a csv reader for 431TT test result written in python
#-------------------------------------------------------------------------------------------------------------
#import tools for python
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import statistics
import logging


logging.basicConfig(filename='csv_analyzer.log', level=logging.WARNING, 
                    format='%(asctime)s - %(levelname)s - %(message)s')


#get file from user-------------------------------------------------------------------------------------------
def get_file_path(): 
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select CSV file",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting program.")
        exit()
    return file_path

#get user input coefficient-----------------------------------------------------------------------------------
def get_coefficient():
    while True:
        coef = simpledialog.askfloat("Input", "Enter the coefficient (1-6):", 
                                     minvalue=1, maxvalue=6)
        if coef is not None:
            return coef
        messagebox.showerror("Error", "Please enter a valid number between 1 and 6.")
     
#get user column choice, parameter can be changed based on program needs--------------------------------------
def get_measurement_number():
    while True:
        measurement = simpledialog.askstring("Input", "Enter the Measurement # (1-7):")
        if measurement is None:
            messagebox.showerror("Error", "Measurement number is required. Exiting program.")
            exit()
        try:
            measurement = int(measurement)
            if 1 <= measurement <= 7:
                return measurement
            else:
                messagebox.showerror("Error", "Please enter a number between 1 and 7.")
        except ValueError:
            messagebox.showerror("Error", "Please enter a valid number.")
       


#read csv provided from user, the csv's 10th row contains the measurement number------------------------------
def read_csv(file_path, measurement_number):
    data = {"values": [], "serial_numbers": []}
    column_name = ""
    try:
        with open(file_path, 'r') as csvfile:
            reader = csv.reader(csvfile)
            
            # Skip to the 'Measure #' row (row 10)
            for _ in range(9):
                next(reader)
            
            # Read the 'Measure #' row to get column names
            measure_row = next(reader)
            column_index = measurement_number  # This is correct now
            if column_index < 1 or column_index >= len(measure_row):
                messagebox.showerror("Error", f"Invalid Measurement # {measurement_number}")
                return None, None
            column_name = measure_row[column_index]
            
            # Skip to the data rows (start from row 28)
            for _ in range(17):
                next(reader)
            
            # Read the data
            for row_index, row in enumerate(reader, start=28):
                try:
                    if len(row) > column_index:
                        if row[column_index].strip() == '':
                            logging.warning(f"Empty cell at row {row_index}, column {column_index}")
                            continue
                        value = float(row[column_index])
                        data["values"].append(value)
                        serial = int(row[0]) if row[0].strip() and row[0].strip().isdigit() else row_index
                        data["serial_numbers"].append(serial)
                    else:
                        logging.warning(f"Row {row_index} does not have enough columns")
                except (ValueError, IndexError) as e:
                    logging.warning(f"Error in row {row_index}, column {column_index}: {str(e)}")
    except Exception as e:
        logging.error(f"An error occurred while reading the CSV file: {str(e)}", exc_info=True)
        messagebox.showerror("Error", f"An error occurred while reading the CSV file. Check the log for details.")
        return None, None
    
    if not data["values"]:
        messagebox.showerror("Error", f"No valid data found in column {column_name} (Measurement #{measurement_number}).")
        return None, None
    
    if len(data["values"]) != len(data["serial_numbers"]):
        logging.warning("Mismatch between number of values and serial numbers")
        messagebox.showwarning("Warning", "Mismatch between number of values and serial numbers. Results may be inconsistent.")
    
    return data, column_name

#data analysis------------------------------------------------------------------------------------------------
def analyze_data(data, coefficient):
    if not data["values"]:
        messagebox.showerror("Error", "No valid data for analysis.")
        return None
    
    try:
        avg = statistics.mean(data["values"])
        std_dev = statistics.stdev(data["values"]) if len(data["values"]) > 1 else 0
        value1 = avg - std_dev * coefficient
        value2 = avg + std_dev * coefficient
        return avg, std_dev, value1, value2
    except statistics.StatisticsError as e:
        messagebox.showerror("Error", f"Statistical error: {str(e)}")
        return None


#output-------------------------------------------------------------------------------------------------------

def get_unique_output_path(base_path):
    directory, filename = os.path.split(base_path)
    name, ext = os.path.splitext(filename)
    counter = 1
    while True:
        new_path = os.path.join(directory, f"{name}_PAT{counter}.csv")
        if not os.path.exists(new_path):
            return new_path
        counter += 1
        

def write_output(output_path, input_path, column_name, avg, std_dev, coefficient, value1, value2, data):
    unique_output_path = get_unique_output_path(output_path)
    try:
        with open(unique_output_path, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow([input_path])
            writer.writerow([column_name])
            writer.writerow([])
            writer.writerow(["Average", f"{avg:.10f}"])
            writer.writerow(["Standard Deviation", f"{std_dev:.10f}"])
            writer.writerow(["Coefficient", f"{coefficient:.2f}"])
            writer.writerow(["PAT Lower Limit", f"{value1:.10f}"])
            writer.writerow(["PAT Higher Limit", f"{value2:.10f}"])
            writer.writerow([])
            writer.writerow(["Outlier data"])
            writer.writerow(["Serial#", "Value"])
            
            for serial, value in zip(data["serial_numbers"], data["values"]):
                if value < value1 or value > value2:
                    writer.writerow([serial, f"{value:.10f}"])

        messagebox.showinfo("Success", f"Analysis complete. Results written to {unique_output_path}")
    except Exception as e:
        logging.error(f"An error occurred while writing the output file: {str(e)}", exc_info=True)
        messagebox.showerror("Error", f"An error occurred while writing the output file. Check the log for details.")

    
#main--------------------------------------------------------------------------------------------------------
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(title="Select CSV file", filetypes=[("CSV files", "*.csv")])
    if not file_path:
        messagebox.showerror("Error", "No file selected. Exiting program.")
        return

    coefficient = simpledialog.askfloat("Input", "Enter the coefficient (1-6):", minvalue=1, maxvalue=6)
    if coefficient is None:
        messagebox.showerror("Error", "Invalid coefficient. Exiting program.")
        return

    measurement_number = simpledialog.askinteger("Input", "Enter the Measurement # (1-7):", minvalue=1, maxvalue=7)
    if measurement_number is None:
        messagebox.showerror("Error", "Invalid Measurement #. Exiting program.")
        return

    data, column_name = read_csv(file_path, measurement_number)
    if data is None:
        return

    result = analyze_data(data, coefficient)
    if result is None:
        return

    avg, std_dev, value1, value2 = result

    output_path = os.path.splitext(file_path)[0] + "_PAT.csv"
    write_output(output_path, file_path, column_name, avg, std_dev, coefficient, value1, value2, data)

if __name__ == "__main__":
    main()

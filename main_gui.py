import tkinter as tk
from tkinter import filedialog
import threading
from geopy.geocoders import Nominatim
from geopy.distance import geodesic
import pandas as pd
import time
import math
import os
import logging

class GUIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Processor")
        
        # File path variable
        self.file_path = tk.StringVar()
        
        # Status variable
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")

        # Threading variable
        self.processing_thread = None

        # GUI Components
        self.label = tk.Label(root, text="Select File:")
        self.label.pack()

        self.entry = tk.Entry(root, textvariable=self.file_path, state='disabled', width=50)
        self.entry.pack(pady=5)

        self.browse_button = tk.Button(root, text="Browse", command=self.browse_file, width=20)
        self.browse_button.pack(pady=5)

        self.run_button = tk.Button(root, text="Run Program", command=self.run_program, width=20)
        self.run_button.pack(pady=5)

        self.status_label = tk.Label(root, textvariable=self.status_var)
        self.status_label.pack()

        # Bind the window close event to the function that stops the processing thread
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.file_path.set(file_path)

    def run_program(self):
        file_path = self.file_path.get()

        if not file_path:
            self.status_var.set("Please select a file.")
            return

        if not file_path.lower().endswith('.xlsx'):
            self.status_var.set("Selected file must be in Excel (.xlsx) format.")
            return

        self.status_var.set("Running...")

        # Use threading to run the processing in the background
        self.processing_thread = threading.Thread(target=self.run_processing, args=(file_path,))
        self.processing_thread.start()

        # Check the thread periodically and update the status
        self.root.after(100, self.check_thread)

    def check_thread(self):
        if self.processing_thread and self.processing_thread.is_alive():
            self.status_var.set("Running...")
            self.root.after(100, self.check_thread)
        else:
            self.status_var.set("Program completed successfully.")

    def run_processing(self, file_path):
        output_file_path = os.path.splitext(file_path)[0] + '_output.xlsx'
        new_headers = ['Customers Name', 'Customers City', 'Customers State', 'Architects Name', 'Architects City', 'Architects State', 
                       'Project Name', 'Project Id', 'Project City', 'Project State',
                       'Project Type', 'Project Facility Type', 'Section ID', 'Details', 'BPM: Manufacturer/Product',
                       'Date: Start Date', 'Date: Last Change', 'Company Name', 'Company Address', 'Company Address2',
                       'Company City', 'Company State', 'Company ZIP', 'Company Country', 'Company Phone1']

        def setup_logger(output_file_path):
            log_file_path = os.path.splitext(output_file_path)[0] + '_log.txt'
            logging.basicConfig(filename=log_file_path, filemode='w', level=logging.INFO)
            return logging.getLogger(__name__)

        logger = setup_logger(output_file_path)

        def log_print(message):
            logger.info(message)
            print(message)

        def get_coordinates(address_array, sheet_name):
            # Initialize the geocoder for Nominatim
            geolocator = Nominatim(user_agent="my_geocoder")

            result_list = []

            for i, address_tuple in enumerate(address_array):
                # Unpack the address tuple
                city, state = address_tuple

                if is_empty(city) and is_empty(state):
                    result_list.append((i, city, state, None))
                    log_print(f"{sheet_name} Row {i}, Empty address. Coordinates set to None.")
                    continue

                if is_empty(city):
                    city = ""

                if is_empty(state):
                    state = ""

                if city.lower() == "undefined":
                    city = ""
                    result_list.append((i, city, state, None))
                    log_print(f"{sheet_name} Row {i}, Empty address. Coordinates set to None.")
                    continue

                if state.lower() == "undefined":
                    state = ""
                    result_list.append((i, city, state, None))
                    log_print(f"{sheet_name} Row {i}, Empty address. Coordinates set to None.")
                    continue

                # Concatenate city and state for accurate geocoding
                location_query = f"{city}, {state}, USA"

                retry_count = 0
                while retry_count < 10:
                    try:
                        # Get the coordinates
                        location = geolocator.geocode(location_query)

                        if location:
                            coordinates = (location.latitude, location.longitude)
                            result_list.append((i, city, state, coordinates))
                            log_print(f"{sheet_name} Row {i}, Address: {city}, {state}, Coordinates: {coordinates}")
                        else:
                            result_list.append((i, city, state, None))
                            log_print(f"{sheet_name} Row {i}, Coordinates not found for {location_query}")
                        break  # Exit the retry loop if successful
                    except Exception as e:
                        log_print(f"--------->Error: {e}")
                        retry_count += 1
                        log_print(f"Retrying... Attempt {retry_count}")
                        time.sleep(1)

            return result_list

        def extract_city_state(file_path, sheet_name, city_column_index, state_column_index):
            # Read the Excel file
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # Create an array of tuples (city, state)
            city_state_array = list(zip(df.iloc[:, city_column_index], df.iloc[:, state_column_index]))

            return city_state_array

        def is_empty(value):
            if pd.isna(value) or (isinstance(value, float) and math.isnan(value)):
                return True
            return False

        def find_closest_locations(array1, array2):
            result_list = []

            for loc1_info in array1:
                loc1 = loc1_info[-1]  # Extract only the coordinates from the array element
                closest_distance = float('inf')
                closest_location_info = None

                for loc2_info in array2:
                    loc2 = loc2_info[-1]  # Extract only the coordinates from the array element
                    distance = geodesic(loc1, loc2).miles

                    if distance < closest_distance:
                        closest_distance = distance
                        closest_location_info = loc2_info + (distance,)

                result_list.append((loc1_info, closest_location_info))

            return result_list

        def select_and_save_rows(input_file_path, sheet_name, row_indices):
            # Load the Excel file into a DataFrame without headers
            df = pd.read_excel(input_file_path, sheet_name=sheet_name, header=None)

            # Create a new DataFrame to store the selected rows
            selected_rows_df = df.iloc[row_indices, :]
            return selected_rows_df

        def concatenate_and_save(df, output_file_path):
            log_print(f"Concatenating and saving DataFrame to '{output_file_path}'.")

            if os.path.exists(output_file_path):
                log_print(f"The output file '{output_file_path}' already exists. Replacing with new data.")
                os.remove(output_file_path)

            df.to_excel(output_file_path, index=False)
            log_print(f"The concatenated DataFrame is saved to '{output_file_path}'.")

        def change_column_headers(input_file_path, new_headers):
            log_print(f"Changing column headers in the Excel file '{input_file_path}'.")

            df = pd.read_excel(input_file_path)
            if len(new_headers) != len(df.columns):
                log_print("Error: The number of new headers does not match the number of columns.")
                return

            df.columns = new_headers
            df.to_excel(input_file_path, index=False)

            log_print(f"The column headers in the Excel file '{input_file_path}' have been updated.")

        # Extract city and state arrays from the "customer" tab (columns B and C)
        customer_array = extract_city_state(file_path, 'Customers', 1, 2)
        log_print("Customers Tab:")
        log_print(customer_array)

        coordinates_customer = get_coordinates(customer_array, "Customers")
        log_print(coordinates_customer)

        architects_array = extract_city_state(file_path, 'Architects', 1, 2)
        log_print("Architects Tab:")
        log_print(architects_array)

        coordinates_architects = get_coordinates(architects_array, "Architects")
        log_print(coordinates_architects)

        specs_array = extract_city_state(file_path, 'Specs', 2, 3)
        log_print("Specs Tab:")
        log_print(specs_array)

        coordinates_specs = get_coordinates(specs_array, "Specs")
        log_print(coordinates_specs)

        customer_row = []
        architect_row = []
        spec_row = []

        closest_locations = find_closest_locations(coordinates_customer, coordinates_architects)
        for item in closest_locations:
            original_location_info, closest_location_info = item
            log_print(f"Customer Index Info: {original_location_info[0]}, Closest Architect Index Info: {closest_location_info[0]}")
            customer_row.append(original_location_info[0])
            architect_row.append(closest_location_info[0])

        closest_locations = find_closest_locations(coordinates_customer, coordinates_specs)
        for item in closest_locations:
            original_location_info, closest_location_info = item
            log_print(f"Customer Index Info: {original_location_info}, Closest Specs Index Info: {closest_location_info}")
            spec_row.append(closest_location_info[0])

        selected_rows_df1 = select_and_save_rows(file_path, "Customers", customer_row)
        log_print(f"Customers Sheet Info:{selected_rows_df1}")
        selected_rows_df2 = select_and_save_rows(file_path, "Architects", architect_row)
        log_print(f"Architects Sheet Info:{selected_rows_df2}")
        selected_rows_df3 = select_and_save_rows(file_path, "Specs", spec_row)
        log_print(f"Specs Sheet Info:{selected_rows_df3}")

        selected_rows_df1 = selected_rows_df1.reset_index(drop=True)
        selected_rows_df2 = selected_rows_df2.reset_index(drop=True)
        selected_rows_df3 = selected_rows_df3.reset_index(drop=True)

        df = pd.concat([selected_rows_df1, selected_rows_df2, selected_rows_df3], axis=1)

        concatenate_and_save(df, output_file_path)
        change_column_headers(output_file_path, new_headers)

    def on_close(self):
        # Stop the processing thread if it is running
        if self.processing_thread and self.processing_thread.is_alive():
            self.processing_thread.join()  # Wait for the thread to finish
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = GUIApp(root)
    root.mainloop()

import pandas as pd
import numpy as np
import json
import csv
import os
from datetime import datetime


class SecretSanta:
    def __init__(self, employee_file, secret_santa_file):
        """
        Initialize the SecretSanta class with file paths for employee and Secret Santa data.
        """
        self.employee_file = employee_file
        self.secret_santa_file = secret_santa_file
        self.employee_df = None
        self.secret_santa_df = None
        self.secret_santa_assignments = []

    def load_data(self):
        """
        Load data from Excel files into pandas DataFrames.
        """
        self.employee_df = pd.read_excel(self.employee_file, engine='openpyxl')
        self.secret_santa_df = pd.read_excel(self.secret_santa_file, engine='openpyxl')

    def assign_secret_santa(self):

        """
        Assign Secret Santa pairs while avoiding duplicates and invalid matches.
        """
        secret_santa_data = np.array(self.secret_santa_df)
        secret_child_data = np.array([
            self.secret_santa_df['Secret_Child_Name'].values,
            self.secret_santa_df['Secret_Child_EmailID'].values
        ])
        assigned_email_ids = []

        for employee_email in self.employee_df['Employee_EmailID']:
            
            employee_index = np.where(self.secret_santa_df['Employee_EmailID'] == employee_email)
            
            if employee_index[0].size > 0: 
                employee_index = employee_index[0][0]
                candidate_count = 0
                
                for candidate_email in secret_child_data[1]:
                    if (
                        candidate_email != secret_santa_data[employee_index][3] and  
                        candidate_email != secret_santa_data[employee_index][1] and  
                        candidate_email not in assigned_email_ids                  
                    ):
                        
                        assigned_email_ids.append(candidate_email)
                        self.secret_santa_assignments.append({
                            'Employee_Name': secret_santa_data[employee_index][0],
                            'Employee_EmailID': secret_santa_data[employee_index][1],
                            'Secret_Child_Name': secret_child_data[0][candidate_count],
                            'Secret_Child_EmailID': candidate_email
                        })
                        break
                    candidate_count += 1


    def save_to_csv(self, output_file):
        """
        Save the Secret Santa assignments to a CSV file.
        """
 
        with open(output_file, mode='w', newline='') as file:
            writer = csv.DictWriter(file, fieldnames=self.secret_santa_assignments[0].keys())
            writer.writeheader()
            writer.writerows(self.secret_santa_assignments)
            

    def save_to_json(self, output_file):
        """
        Save the Secret Santa assignments to a JSON file.
        """
        with open(output_file, mode='w') as file:
            json.dump(self.secret_santa_assignments, file, indent=4)


if __name__ == "__main__":
    current_directory = os.getcwd()

    # Define child folder path
    child_path = "/Employeedetails"
    result_path = "/Result"

    # Combine paths
    full_path = current_directory + child_path
    result = current_directory + result_path

    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    outputcsv = f"secret_santa_{timestamp}.csv"
    outputjson = f"secret_santa_{timestamp}.json"
    # Hardcode the file names
    file1 = "EmployeeList.xlsx"
    file2 = "Secret-Santa-Game-Result-2023.xlsx"
    file3 = outputcsv
    file4 = outputjson

    # Define full file paths
    try:
        if not os.path.exists(result):
           os.makedirs(result)
        employee_file_path = os.path.join(full_path, file1)
        secret_santa_file_path = os.path.join(full_path, file2)
        output_csv_file = os.path.join(result, file3)
        output_json_file = os.path.join(result, file4)
        if not os.path.exists(employee_file_path):
            raise FileNotFoundError(f"Employee file not found: {employee_file_path}")
        if not os.path.exists(secret_santa_file_path):
            raise FileNotFoundError(f"Secret Santa file not found: {secret_santa_file_path}")

        secret_santa = SecretSanta(employee_file_path, secret_santa_file_path)
        secret_santa.load_data()
        secret_santa.assign_secret_santa()
        secret_santa.save_to_csv(output_csv_file)
        secret_santa.save_to_json(output_json_file)
        print(f"The file has been successfully finished. Storage location: Result/{file3}")
        print(f"The file has been successfully finished. Storage location:: Result/{file4}")
    except FileNotFoundError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

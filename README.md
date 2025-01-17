# DigitalXc
**Secret Santa Assignment Program**

**Overview**

This Python program automates the process of assigning Secret Santa pairs within an organization. It uses employee data from an Excel sheet and generates a list of Secret Santa assignments while ensuring no one is assigned to themselves and avoiding duplicate pairings. The results are saved both as a CSV file and a JSON file.

**Features**

* Reads employee and Secret Santa data from Excel files.
* Assigns Secret Santa pairs avoiding duplicates and invalid matches.
* Outputs results to both CSV and JSON formats for easy sharing and use.
* Handles file reading and writing with error handling.

**Requirements**

Before running the program, make sure the following dependencies are installed:

* Python 3.x
* Pandas
* Numpy
* Openpyxl (for reading Excel files)

You can install the required dependencies using pip:

* pip install pandas numpy openpyxl

**File Structure**

The program requires two input files:

* EmployeeList.xlsx: Contains the list of employees with their names and email IDs.
* Secret-Santa-Game-Result-2023.xlsx: Contains data on who has been assigned to whom (existing Secret Santa data).

The program generates two output files:

* secret_santa_assignments.csv: A CSV file containing the final Secret Santa pairings.
* secret_santa_assignments.json: A JSON file containing the final Secret Santa pairings in a structured format.


**Instructions for Running the Program**

Step 1: Prepare Input Files

**Ensure you have the following files:**

    * EmployeeList.xlsx: A spreadsheet that contains employee information (names and email IDs).
    * Secret-Santa-Game-Result-2023.xlsx: A spreadsheet that contains existing Secret Santa data.

**The Excel files should have the following structure:**

**EmployeeList.xlsx should have columns like:**

    * Employee_Name, Employee_EmailID

    Secret-Santa-Game-Result-2023.xlsx should have columns like:

    * Employee_Name, Employee_EmailID, Secret_Santa_Name, Secret_Child_Name, Secret_Child_EmailID

Step 2: Place Files in Correct Directories

    The input files should be placed in the /Employeedetails folder, and the program will save the output in the /Result folder.

Step 3: Run the Program
    You can execute the program by running the following command in your terminal:

    python employee.py

This will:

    * Load the input data from the Excel files.

    * Assign Secret Santa pairs.

    * Save the results to secret_santa_assignments.csv and secret_santa_assignments.json in the /Result folder.

Step 4: Check the Results

After running the program, you can find the results in the /Result directory:

secret_santa_assignments.csv: Contains the employee name, email, and the assigned Secret Santa pair.

secret_santa_assignments.json: A JSON representation of the same data.

**Error Handling**

If the input files are not found in the expected directory, the program will raise a FileNotFoundError with the respective error message.

Any unexpected errors during the execution will be caught and displayed as a general exception.
Example Output (CSV)

**The secret_santa_assignments.csv file will look something like this:**

Employee_Name,Employee_EmailID,Secret_Child_Name,Secret_Child_EmailID

John Doe,john.doe@example.com,Jane Smith,jane.smith@example.com

Alice Brown,alice.brown@example.com,Bob White,bob.white@example.com
...

**Example Output (JSON)**

**The secret_santa_assignments.json file will look like this:**

[
    {
        "Employee_Name": "John Doe",
        "Employee_EmailID": "john.doe@example.com",
        "Secret_Child_Name": "Jane Smith",
        "Secret_Child_EmailID": "jane.smith@example.com"
    },
    {
        "Employee_Name": "Alice Brown",
        "Employee_EmailID": "alice.brown@example.com",
        "Secret_Child_Name": "Bob White",
        "Secret_Child_EmailID": "bob.white@example.com"
    }
]



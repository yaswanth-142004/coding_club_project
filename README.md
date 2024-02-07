# Contact Management System

This is a simple Contact Management System implemented in Python using Tkinter for the GUI and OpenPyXL for handling Excel files. The system allows users to add, search, update, and delete contacts from an Excel-based database.

## Features

- *Add Contact*: Users can add new contacts to the database by providing the necessary details such as name, mobile number, email ID, registration number, and course enrolled.
  
- *Search Contact*: Users can search for a contact by name and view the details if found.

- *Update Contact*: Users can update the details of an existing contact such as mobile number, email ID, registration number, or course enrolled.

- *Delete Contact*: Users can delete a contact from the database.

## Requirements

- Python 3.x
- openpyxl (install via pip: pip install openpyxl)

## How to Use

1. Clone the repository to your local machine:

    bash
    git clone https://github.com/yourusername/contact-management-system.git
    

2. Navigate to the project directory:

    bash
    cd contact-management-system
    

3. Ensure you have Python installed. Install the required dependencies:

    bash
    pip install openpyxl
    

4. Run the application:

    bash
    python contact_management.py
    

5. The Contact Management System GUI will appear. You can now use the various features to manage your contacts.

## Note

- The database is stored in an Excel file named dbms.xlsx. Ensure this file exists in the project directory before running the application.

- Make sure to provide valid inputs as per the specified format when adding or updating contacts.

- For any issues or suggestions, feel free to open an issue in the GitHub repository.

Happy Contact Managing!

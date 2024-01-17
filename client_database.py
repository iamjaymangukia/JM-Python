import tkinter as tk
from tkinter import messagebox
import pyodbc

class EmployeeDataEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Client Contact Data Entry")
        self.root.geometry("600x500")  # Set the size of the GUI window

        # Configure column and row weights to center the content
        for i in range(5):
            self.root.columnconfigure(i, weight=1)
            self.root.rowconfigure(i, weight=1)
        self.root.rowconfigure(5, weight=1)

        # Employee details variables
        self.company_name_var = tk.StringVar()
        self.client_name_var = tk.StringVar()
        self.designation_var = tk.StringVar()
        self.department_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.office_address_var = tk.StringVar()
        self.landline_number_var = tk.StringVar()
        self.mobile_number_var = tk.StringVar()
        self.company_website_var = tk.StringVar()
        self.contact_responsible_person = tk.StringVar()
        self.remarks_var = tk.StringVar()

        # Entry widgets and labels
        tk.Label(root, text="Company Name:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.company_name_var, width=40).grid(row=0, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Client Name:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
        client_entry = tk.Entry(root, textvariable=self.client_name_var, width=40)
        client_entry.grid(row=1, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Designation:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.designation_var, width=40).grid(row=2, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Department:").grid(row=3, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.department_var, width=40).grid(row=3, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Email ID:").grid(row=4, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.email_var, width=40).grid(row=4, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Office Address:").grid(row=5, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.office_address_var, width=40).grid(row=5, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Landline Number:").grid(row=6, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.landline_number_var, width=40).grid(row=6, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Mobile Number:").grid(row=7, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.mobile_number_var, width=40).grid(row=7, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Company Website:").grid(row=8, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.company_website_var, width=40).grid(row=8, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Contact Responsible Person:").grid(row=8, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.contact_responsible_person, width=40).grid(row=8, column=1, padx=10, pady=10, sticky=tk.W)

        tk.Label(root, text="Remarks:").grid(row=9, column=0, padx=10, pady=10, sticky=tk.W)
        tk.Entry(root, textvariable=self.remarks_var, width=40).grid(row=9, column=1, padx=10, pady=10, sticky=tk.W)

        # Submit button
        tk.Button(root, text="Submit", command=self.submit_data).grid(row=10, column=0, columnspan=2, pady=20)

        # Close button event
        root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Duplicacy check when client name is being written
        client_entry.bind("<FocusOut>", self.check_duplicate_entry_on_focus_out)

    def check_duplicate_entry_on_focus_out(self, event):
        # Check for duplicacy when client name is being written
        company_name = self.company_name_var.get()
        client_name = self.client_name_var.get()

        if company_name and client_name:
            if self.check_duplicate_entry(company_name, client_name):
                messagebox.showerror("Error", "Duplicate entry. Client with the same company name and client name already exists.")

    def submit_data(self):
        # Get data from entry widgets
        company_name = self.company_name_var.get()
        client_name = self.client_name_var.get()
        designation = self.designation_var.get()
        department = self.department_var.get()
        email = self.email_var.get()
        office_address = self.office_address_var.get()
        landline_number = self.landline_number_var.get()
        mobile_number = self.mobile_number_var.get()
        company_website = self.company_website_var.get()
        contact_responsible_person = self.contact_responsible_person.get()
        remarks = self.remarks_var.get()

        # Validate mandatory fields
        if not company_name or not client_name or not email:
            messagebox.showerror("Error", "Company Name, Client Name, and Email ID are mandatory fields.")
            return

        # Check for duplicacy before inserting data into the SQL Server database
        if not self.check_duplicate_entry(company_name, client_name):
            # Insert data into the SQL Server database
            self.insert_data_into_database(company_name, client_name, designation, department, email, office_address, landline_number, mobile_number, company_website, contact_responsible_person, remarks)

            # Show a success message
            messagebox.showinfo("Success", "Client data submitted successfully!")

            # Clear entry fields after submission
            self.clear_entry_fields()
        else:
            # Show an error message for duplicacy
            messagebox.showerror("Error", "Duplicate entry. Client with the same company name and client name already exists.")

    def check_duplicate_entry(self, company_name, client_name):
        try:
            # Connect to the SQL Server database
            connection_string = "Driver={SQL Server};Server=JAY\SQLEXPRESS;Database=jmdb;Trusted_Connection=True;"
            connection = pyodbc.connect(connection_string)

            # Create a cursor object
            cursor = connection.cursor()

            # Check if there is a duplicate entry with the same company name and client name
            cursor.execute("SELECT COUNT(*) FROM Employee WHERE Company_Name = ? AND Client_Name = ?", company_name, client_name)
            count = cursor.fetchone()[0]

            # Close the cursor and connection
            cursor.close()
            connection.close()

            return count > 0

        except Exception as e:
            # Handle exceptions (e.g., database connection error)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            return False

    def insert_data_into_database(self, company_name, client_name, designation, department, email, office_address, landline_number, mobile_number, company_website, contact_responsible_person, remarks):
        try:
            # Connect to the SQL Server database
            connection_string = "Driver={SQL Server};Server=JAY\SQLEXPRESS;Database=jmdb;Trusted_Connection=True;"
            connection = pyodbc.connect(connection_string)

            # Create a cursor object
            cursor = connection.cursor()

            # Insert data into the Employee table
            cursor.execute("INSERT INTO Employee (Company_Name, Client_Name, Designation, Department, Email_ID, Office_Address, LandLine_Number, Mobile_Number, Company_Website, Contact_Responsible_Person, Remarks) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                           (company_name, client_name, designation, department, email, office_address, landline_number, mobile_number, company_website, contact_responsible_person, remarks))

            # Commit the transaction
            connection.commit()

            # Close the cursor and connection
            cursor.close()
            connection.close()

        except Exception as e:
            # Handle exceptions (e.g., database connection error)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def clear_entry_fields(self):
        # Clear the entry fields after submission
        self.company_name_var.set("")
        self.client_name_var.set("")
        self.designation_var.set("")
        self.department_var.set("")
        self.email_var.set("")
        self.office_address_var.set("")
        self.landline_number_var.set("")
        self.mobile_number_var.set("")
        self.company_website_var.set("")
        self.contact_responsible_person.set("")
        self.remarks_var.set("")

    def on_close(self):
        # Close the application when clicking the close button
        self.root.destroy()

# Create the main application window
root = tk.Tk()
app = EmployeeDataEntryApp(root)

# Start the main event loop
root.mainloop()

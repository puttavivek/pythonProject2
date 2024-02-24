import pandas as pd
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
from math import ceil


class Inventory:
    """
        Class for managing inventory data, including operations like adding, updating, searching, and printing.
    """
    def __init__(self, filename = "inventory.xlsx"):
        """
                Initializes the Inventory class with the provided or default Excel file.

                Parameters:
                - filename (str): The name of the Excel file containing inventory data.
        """
        self.filename = filename
        self.date_now = datetime.datetime.now()
        self.date_today = self.date_now.strftime("%d-%m-%Y %H:%M")
        self.df = pd.read_excel(self.filename)

    def get_list(self, part_no):
        """
                Retrieves a list of inventory items based on the provided part number.

                Parameters:
                - part_no (str): The part number to search for.

                Returns:
                - pandas.DataFrame: DataFrame containing the selected rows.
        """
        selected_row = self.df.loc[self.df['Part_No'] == part_no]
        return selected_row

    def add_method(self, *args):
        """
                Adds new inventory data or updates existing data based on the provided arguments.

                Parameters:
                - args: Variable arguments containing inventory data.

                Returns:
                - str: Confirmation message indicating whether the operation was successful.
        """
        selected_row = self.get_list(args[1])

        if selected_row.empty:

            input_list = [arg for arg in args]


            input_list.append(self.date_today)
            input_list = [input_list]

            new_data = pd.DataFrame(input_list, columns=self.df.columns)

            updated_df = pd.concat([self.df, new_data], ignore_index=True)

            self.df = updated_df

            try:
                self.df.to_excel(self.filename, index=False)
            except PermissionError:
                return "Please close the excel file that is opened"

            return "Added"

        else:
            self.df.at[selected_row.index[0], 'Quantity'] += args[1]

            try:
                self.df.to_excel(self.filename, index=False)
            except PermissionError:
                return "Please close the excel file that is opened"

            return "Updated"

    def update_method(self, part_no, quantity):
        """
                Updates the quantity of an item in the inventory based on the provided part number.

                Parameters:
                - part_no (str): The part number of the item to update.
                - quantity (int): The quantity to subtract from the available quantity.

                Returns:
                - str: Confirmation message indicating whether the operation was successful or an error occurred.
        """
        part_no = part_no.strip()
        selected_row = self.get_list(part_no)

        if selected_row.empty:
            error = f"No rows found for Part_No = {part_no}"
            return error

        for index, row in selected_row.iterrows():
            if row['Quantity'] >= quantity:
                # Subtract the quantity from the available quantity
                self.df.at[index, 'Quantity'] -= quantity
                try:
                    self.df.to_excel(self.filename, index=False)
                except PermissionError:
                    return "Please close the excel file that is opened"
                return "updated"
            else:
                error = f"Quantity is more than available, Available Quantity is - {row['Quantity']}"
                return error

    def search_method(self, part_no):
        """
                Searches for an item in the inventory based on the provided part number.

                Parameters:
                - part_no (str): The part number to search for.

                Returns:
                - pandas.DataFrame: DataFrame containing the selected rows or an error message.
        """
        part_no = part_no.strip()
        selected_row = self.get_list(part_no)

        if selected_row.empty:
            error = f"No rows found for Part_No = {part_no}"
            return error
        else:
            return selected_row

    def bring_list(self):
        """
                Retrieves the entire inventory list.

                Returns:
                - pandas.DataFrame: DataFrame containing the entire inventory.
        """
        return self.df

    def bulk_entry(self, filename_new):
        """
               Adds bulk inventory data from a file to the existing inventory.

               Parameters:
               - filename_new (str): The name of the file containing bulk inventory data.

               Returns:
               - str: Confirmation message indicating whether the operation was successful.
        """
        new_data = pd.read_excel(filename_new)
        updated_df = pd.concat([self.df, new_data], ignore_index=True)

        self.df = updated_df
        try:
            self.df.to_excel(self.filename, index=False)
        except PermissionError:
            return "Please close the excel file that is opened"

        return "Added"

    def print_list(self, pdf_file_location="."):
        """
                Prints the inventory data to a PDF file.

                Parameters:
                - pdf_file_location (str): The location to save the PDF file.

                Returns:
                - str: Confirmation message indicating the location and name of the printed PDF file.
                """

        max_rows_per_page = 50
        num_pages = len(self.df) / max_rows_per_page
        num_pages = ceil(num_pages)

        relative_column_widths = [0.3, 0.15, 0.1, 0.075, 0.075, 0.3]

        pdf_file_name = f"inventory-{self.date_now.strftime('%d-%m-%Y-%H-%M')}.pdf"
        pdf_file_path = f"{pdf_file_location}/{pdf_file_name}"

        with PdfPages(pdf_file_path) as pdf:
            for page in range(num_pages):
                fig, ax = plt.subplots(figsize=(8.27, 11.69))
                ax.axis('off')  # Turn off axis labels and ticks

                if page == 0:
                    # Add title to the first page
                    fig.text(0.5, 0.97, 'Inventory List', fontsize=16, fontweight='bold', ha='center', va='top')
                    # Adjusting the spacing to reduce the gap between title and table
                    fig.subplots_adjust(top=0.99)

                # Calculate the rows to display on this page
                start_row = page * max_rows_per_page
                end_row = min(start_row + max_rows_per_page, len(self.df))
                df_chunk = self.df.iloc[start_row:end_row]

                # Create the table and add it to the axes
                table = ax.table(cellText=df_chunk.values, colLabels=self.df.columns, cellLoc='center', loc='center', colWidths=relative_column_widths)
                table.auto_set_font_size(False)
                table.set_fontsize(6)

                table.scale(1, 1.2)
                ax.set_position([0, 0, 1, 0.9])

                pdf.savefig(fig, bbox_inches='tight')
                plt.close()

        message = f"List Printed in the selected location with name - {pdf_file_name}"
        return message


class MyApp:
    """
        Main application class for Inventory Management.
    """


    def __init__(self, root):
        """
                Initialize the application.

                Parameters:
                - root: Tkinter root window
        """
        self.root = root
        self.root.title("Inventory Management")
        self.root.geometry("400x350")
        self.root.iconbitmap('Icon.ico')
        # List to store items
        self.item_list = []
        # set minimum window size value
        root.minsize(400, 350)
        # set maximum window size value
        root.maxsize(400, 350)
        # Create and set up the GUI components
        self.create_widgets()

    def create_widgets(self):
        """
                Create GUI components.
        """
        tk.Label(self.root, text="", font=("System", 2)).pack()
        # Add Button
        tk.Button(self.root, text="Add Inventory", command=lambda: self.open_window("Add Item", ['Part Name', 'Part No', 'Model', 'Stock Location', 'Quantity'], self.add_item), padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()
        # Remove Button
        tk.Button(self.root, text="Edit Inventory", command=lambda: self.open_window("Remove Item", ['Part No', 'Quantity'], self.remove_item), padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()
        # Bring List Button
        tk.Button(self.root, text="Search Inventory", command=lambda: self.open_window("Bring List", ['Part No'], self.search_item ), padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()
        # Bring List Button
        tk.Button(self.root, text="View Inventory", command=self.bring_list, padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()
        # Print List Button
        tk.Button(self.root, text="Print Inventory", command=self.print_list, padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()

        tk.Button(self.root, text="Bulk Input", command=self.file_name_bulk, padx=45, pady=2, font=("System", 8), width=6).pack()
        tk.Label(self.root, text="", font=("System", 2)).pack()

    def open_window(self, title, fields, command_function):
        """
                Open a new window for user input.

                Parameters:
                - title: Title of the window
                - fields: List of input fields
                - command_function: Function to execute on submission
        """
        generic_window = tk.Toplevel(self.root)
        generic_window.title(title)
        generic_window.geometry("400x350")
        generic_window.iconbitmap('Icon.ico')
        self.setup_input_fields(generic_window, fields, command_function)

    def setup_input_fields(self, window, fields, command_function):
        """
                Set up input fields in the window.

                Parameters:
                - window: Tkinter window
                - fields: List of input fields
                - command_function: Function to execute on submission
        """
        for field in fields:
            tk.Label(window, text=f"{field}:").pack()
            entry_var = tk.StringVar()
            tk.Entry(window, textvariable=entry_var, name=f'{field.lower()}_entry').pack()

        tk.Label(window, text="", font=("System", 2)).pack()
        tk.Button(window, text="Submit", command=lambda: command_function(window, fields)).pack()

    def to_list(self, window, fields, Func):
        """
                Convert input fields to a list.

                Parameters:
                - window: Tkinter window
                - fields: List of input fields
                - func: Function identifier
        """
        data = []
        error_occured = False
        for field in fields:

            if field.lower() == 'part name' or field.lower() == 'part no' or field.lower() == 'model':
                try:
                    field_input = str(window.children[field.lower() + '_entry'].get())
                    data.append(field_input)
                except ValueError:
                    self.show_message("Invalid data type. Please enter valid data types for each field.", "Error")
                    error_occured = True

            else:
                try:
                    field_input = int(window.children[field.lower() + '_entry'].get())
                    data.append(field_input)
                except ValueError:
                    self.show_message("Invalid data type. Please enter valid data types for each field.", "Error")
                    error_occured = True

        if not error_occured and all(data):
            data.insert(0, Func)
            self.item_list = data
            window.destroy()
            self.root.destroy()

    def add_item(self, window, fields):
        """
                Add item to the list.

                Parameters:
                - window: Tkinter window
                - fields: List of input fields
        """
        self.to_list(window, fields, "ADD")

    def remove_item(self, window, fields):
        """
                Remove item from the list.

                Parameters:
                - window: Tkinter window
                - fields: List of input fields
        """
        self.to_list(window, fields, "SEARCH")

    def search_item(self, window, fields):
        """
                Search for an item in the list.

                Parameters:
                - window: Tkinter window
                - fields: List of input fields
        """
        self.to_list(window, fields, "REMOVE")

    def bring_list(self):
        """
                Display the list of items in the inventory.
        """
        outlist = ["BRING"]
        self.item_list = outlist
        self.root.destroy()

    def print_list(self):
        """
                Print the list of items.
        """
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            outlist = ["PRINT",folder_selected]
            self.item_list = outlist
            self.root.destroy()
        else:
            return ("Please select a folder")

    def file_name_bulk(self):
        """
                Get the file name for bulk input.
        """
        file_selected = filedialog.askopenfilename()
        self.item_list = ["BULK", file_selected]
        self.root.destroy()

    def get_list(self):
        """
                Get the current list of items.

                Returns:
                - List: Current list of items
        """
        if self.item_list:
            return self.item_list

    def show_message(self,message, message_type):
        """
                Display a message box.

                Parameters:
                - message: Message to display
                - message_type: Type of the message (Success, Error, etc.)
        """
        messagebox.showinfo(message_type, message)

    def display_excel_data(self,df,height):
        """
                Display Excel data in a new window.

                Parameters:
                - df: DataFrame containing data to display
                - height: Height of the Treeview widget
        """
        # Create a new Tkinter window
        window = tk.Tk()
        window.title("Data Display")
        window.iconbitmap('Icon.ico')
        # set minimum window size value
        window.minsize(1000, 2000)

        # set maximum window size value
        window.maxsize(1000, 2000)

        # Create a treeview widget for tabular display
        tree = ttk.Treeview(window, columns=list(df.columns), show="headings",  height=height)
        tree.pack(padx=10, pady=10)

        # Insert column headers
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.CENTER)  # Adjust width as needed

        # Insert DataFrame content into the treeview
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        # Run the Tkinter event loop
        window.mainloop()
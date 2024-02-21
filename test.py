import pandas as pd
import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from tkinter import scrolledtext, ttk
from math import ceil

"""
This class is used for adding, updating, removing, displaying, and printing 
the inventory data that is stored in the excel.
"""


class Inventory:

       # Excel file which is used for doing all these functions

    def __init__(self, filename = "inventory.xlsx"):
        self.filename = filename
        self.date_now = datetime.datetime.now()
        self.date_today = self.date_now.strftime("%d-%m-%Y %H:%M")
        self.df = pd.read_excel(self.filename)

    def get_list(self, part_no):
        selected_row = self.df.loc[self.df['Part_No'] == part_no]
        return selected_row

    def add_method(self, part_name, part_no, model, stock_location, quantity):
        selected_row = self.get_list(part_no)

        if selected_row.empty:
            input_list = [[part_name, model, part_no, stock_location, quantity, self.date_today]]

            new_data = pd.DataFrame(input_list, columns=self.df.columns)

            updated_df = pd.concat([self.df, new_data], ignore_index=True)

            self.df = updated_df

            try:
                self.df.to_excel(self.filename, index=False)
            except PermissionError:
                return "Please close the excel file that is opened"

            return "Added"

        else:
            self.df.at[selected_row.index[0], 'Quantity'] += part_no

            try:
                self.df.to_excel(self.filename, index=False)
            except PermissionError:
                return "Please close the excel file that is opened"

            return "Updated"

    def update_method(self, part_no, quantity):
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
        part_no = part_no.strip()
        selected_row = self.get_list(part_no)

        if selected_row.empty:
            error = f"No rows found for Part_No = {part_no}"
            return error
        else:
            return selected_row

    def bring_list(self):
        return self.df

    def print_list(self, pdf_file_location="."):
        max_rows_per_page = 50  # This is an example, adjust the number as needed
        num_pages = len(self.df) / max_rows_per_page
        num_pages = ceil(num_pages)

        relative_column_widths = [0.3, 0.15, 0.075, 0.1, 0.075, 0.3]

        pdf_file_name = f"inventory-{self.date_now.strftime('%d-%m-%Y-%H-%M')}.pdf"
        pdf_file_path = f"{pdf_file_location}/{pdf_file_name}"

        with PdfPages(pdf_file_path) as pdf:
            for page in range(num_pages):
                fig, ax = plt.subplots(figsize=(8.27, 11.69))
                ax.axis('off')  # Turn off axis labels and ticks

                if page == 0:
                    fig.text(0.5, 0.97, 'Inventory List', fontsize=16, fontweight='bold', ha='center', va='top')
                    # Adjust the spacing to reduce the gap between title and table
                    fig.subplots_adjust(top=0.99)

                # Calculate the rows to display on this page
                start_row = page * max_rows_per_page
                end_row = min(start_row + max_rows_per_page, len(self.df))
                df_chunk = self.df.iloc[start_row:end_row]

                # Create the table and add it to the axes
                table = ax.table(cellText=df_chunk.values, colLabels=self.df.columns, cellLoc='center', loc='center', colWidths=relative_column_widths)
                table.auto_set_font_size(False)
                table.set_fontsize(6)  # Adjust font size as needed

                table.scale(1, 1.2)
                ax.set_position([0, 0, 1, 0.9])

                pdf.savefig(fig, bbox_inches='tight')
                plt.close()

        message = f"List Printed in the selected location with name - {pdf_file_name}"
        return message


class MyApp:
    def __init__(self, root):
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

    def open_window(self, title, fields, command_function):
        generic_window = tk.Toplevel(self.root)
        generic_window.title(title)
        generic_window.geometry("400x350")
        generic_window.iconbitmap('Icon.ico')
        self.setup_input_fields(generic_window, fields, command_function)

    def setup_input_fields(self, window, fields, command_function):
        for field in fields:
            tk.Label(window, text=f"{field}:").pack()
            entry_var = tk.StringVar()
            tk.Entry(window, textvariable=entry_var, name=f'{field.lower()}_entry').pack()

        tk.Label(window, text="", font=("System", 2)).pack()
        tk.Button(window, text="Submit", command=lambda: command_function(window, fields)).pack()

    def add_item(self, window, fields):
        try:
            # Convert input data to the desired data types
            part_name = str(window.children['part name_entry'].get()).strip()
            part_no = str(window.children['part no_entry'].get()).strip()
            model = str(window.children['model_entry'].get()).strip()
            stock_location = int(window.children['stock location_entry'].get())
            quantity = int(window.children['quantity_entry'].get())

            # Check if all fields are filled
            if all((part_name, part_no, stock_location, quantity)):
                data = ["ADD", part_name, part_no, model, stock_location, quantity]
                self.item_list = data
                messagebox.showinfo("Success", "Item added successfully!")
                window.destroy()
                self.root.destroy()
            else:
                messagebox.showwarning("Error", "Please fill in all fields.")

        except ValueError:
            self.show_error_message("Invalid data type. Please enter valid data types for each field."
                                    " Part Name, Part No, Model, must be 'Alphanumerical', Stock Location is a number between 1 to 100"
                                    " and Quantity must be a numerical value")

    def remove_item(self, window, fields):
        data = []
        error_occured = False
        for field in fields:

            if field.lower() == 'part name' or field.lower() == 'part no' or field.lower() == 'model':
                try:
                    field_input = str(window.children[field.lower() + '_entry'].get())
                    data.append(field_input)
                except ValueError:
                    self.show_error_message("Invalid data type. Please enter valid data types for each field.")
                    error_occured = True

            else:
                try:
                    field_input = int(window.children[field.lower() + '_entry'].get())
                    data.append(field_input)
                except ValueError:
                    self.show_error_message("Invalid data type. Please enter valid data types for each field.")
                    error_occured = True

        if not error_occured and all(data):
            data.insert(0, "REMOVE")
            self.item_list = data
            window.destroy()
            self.root.destroy()

    def search_item(self, window, fields):
        data = [window.children[field.lower() + '_entry'].get() for field in fields]
        if all(data):
            data.insert(0, "SEARCH")
            self.item_list = data
            window.destroy()
            self.root.destroy()
        else:
            messagebox.showwarning("Error", "Please fill in all fields.")

    def bring_list(self):
        outlist = ["BRING"]
        self.item_list = outlist
        self.root.destroy()

    def print_list(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            outlist = ["PRINT",folder_selected]
            self.item_list = outlist
            self.root.destroy()
        else:
            self.show_error_message("Please select a folder")

    def get_list(self):
        if self.item_list:
            return self.item_list

    def show_error_message(self, error_message):
        messagebox.showerror("Error", error_message)

    def show_message(self,message):
        # Display the message box
        messagebox.showinfo("Message", message)

    def display_excel_data(self,df,height):
        # Create a new Tkinter window
        window = tk.Tk()
        window.title("Data Display")
        window.iconbitmap('Icon.ico')

        # Create a treeview widget for tabular display
        tree = ttk.Treeview(window, columns=list(df.columns), show="headings",  height=height)
        tree.pack(padx=10, pady=10)

        # Insert column headers
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor=tk.CENTER)  # Adjust width as needed

        # Insert DataFrame content into the treeview
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        # Run the Tkinter event loop
        window.mainloop()
        # # Create a new Tkinter window
        # window = tk.Tk()
        # window.title("Data Display")
        # window.iconbitmap('Icon.ico')
        #
        # # Create a scrolled text widget
        # text_widget = scrolledtext.ScrolledText(window, width=150, height=height, wrap=tk.WORD,  font=("System", 6))
        # text_widget.pack(padx=10, pady=10)
        #
        # # Insert the DataFrame content into the text widget
        # text_widget.insert(tk.INSERT, df.to_string(index=False))
        #
        # # Run the Tkinter event loop
        # window.mainloop()

root = tk.Tk()
app = MyApp(root)
root.mainloop()

gui_input = app.get_list()

if gui_input:
    inventory = Inventory()

    action = gui_input[0]

    if action == "ADD":
        inventory.add_method(*gui_input[1:])
    elif action == "REMOVE":
        return_str = inventory.update_method(*gui_input[1:3])
        app.show_message(return_str) if return_str == "Updated" else app.show_error_message(return_str)
    elif action == "BRING":
        excel_df = inventory.bring_list()
        app.display_excel_data(excel_df, 40)
    elif action == "SEARCH":
        result = inventory.search_method(gui_input[1])
        app.show_message(result) if isinstance(result, str) else app.display_excel_data(result, 20)
    elif action == 'PRINT':
        message = inventory.print_list(gui_input[1])
        app.show_message(message)
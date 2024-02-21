# import pandas as pd
# import datetime
# import matplotlib.pyplot as plt
# from matplotlib.backends.backend_pdf import PdfPages
# import tkinter as tk
# from tkinter import messagebox
# from tkinter import filedialog
# from tkinter import scrolledtext
# from math import ceil
#
# class Inventory:
#     """
#     This class handles adding, updating, removing, displaying, and printing
#     inventory data stored in an Excel file.
#     """
#
#     filename = "inventory.xlsx"  # Excel file used for all operations
#
#     def __init__(self):
#         """
#         Initializes the Inventory object by reading data from the Excel file.
#         """
#         self.date_now = datetime.datetime.now()
#         self.date_today = self.date_now.strftime("%d-%m-%Y %H:%M")
#         self.df = pd.read_excel(self.filename)
#
#     def get_list(self, part_no):
#         """
#         Retrieves a selected row from the DataFrame based on the given part number.
#
#         Args:
#             part_no (str): Part number to search for.
#
#         Returns:
#             pd.Series: Selected row from the DataFrame.
#         """
#         selected_row = self.df.loc[self.df['Part_No'] == part_no]
#         return selected_row
#
#     def add_method(self, part_name, part_no, model, stock_location, quantity):
#         """
#         Adds an item to the inventory or updates the quantity if the item already exists.
#
#         Args:
#             part_name (str): Part name.
#             part_no (str): Part number.
#             model (str): Model information.
#             stock_location (int): Stock location.
#             quantity (int): Quantity to add.
#
#         Returns:
#             str: Status message indicating success or an error.
#         """
#         selected_row = self.get_list(part_no)
#
#         if selected_row.empty:
#             input_list = [[part_name, model, part_no, stock_location, quantity, self.date_today]]
#             new_data = pd.DataFrame(input_list, columns=self.df.columns)
#             updated_df = pd.concat([self.df, new_data], ignore_index=True)
#             self.df = updated_df
#
#             try:
#                 self.df.to_excel(self.filename, index=False)
#             except PermissionError:
#                 return "Please close the Excel file that is opened"
#
#             return "Added"
#
#         else:
#             self.df.at[selected_row.index[0], 'Quantity'] += part_no
#
#             try:
#                 self.df.to_excel(self.filename, index=False)
#             except PermissionError:
#                 return "Please close the Excel file that is opened"
#
#             return "Updated"
#
#     def update_method(self, part_no, quantity):
#         """
#         Updates the inventory by subtracting the specified quantity from the given part number.
#
#         Args:
#             part_no (str): Part number.
#             quantity (int): Quantity to subtract.
#
#         Returns:
#             str: Status message indicating success or an error.
#         """
#         part_no = part_no.strip()
#         selected_row = self.get_list(part_no)
#
#         if selected_row.empty:
#             error = f"No rows found for Part_No = {part_no}"
#             return error
#
#         for index, row in selected_row.iterrows():
#             if row['Quantity'] >= quantity:
#                 self.df.at[index, 'Quantity'] -= quantity
#                 try:
#                     self.df.to_excel(self.filename, index=False)
#                 except PermissionError:
#                     return "Please close the Excel file that is opened"
#                 return "Updated"
#             else:
#                 error = f"Quantity is more than available, Available Quantity is - {row['Quantity']}"
#                 return error
#
#     def search_method(self, part_no):
#         """
#         Searches for a part in the inventory based on the given part number.
#
#         Args:
#             part_no (str): Part number.
#
#         Returns:
#             pd.DataFrame or str: Found rows or an error message.
#         """
#         part_no = part_no.strip()
#         selected_row = self.get_list(part_no)
#
#         if selected_row.empty:
#             error = f"No rows found for Part_No = {part_no}"
#             return error
#         else:
#             return selected_row
#
#     def bring_list(self):
#         """
#         Retrieves the entire inventory.
#
#         Returns:
#             pd.DataFrame: Full inventory.
#         """
#         return self.df
#
#     def print_list(self, pdf_file_location="."):
#         """
#         Prints the inventory to a PDF file.
#
#         Args:
#             pdf_file_location (str): Location to save the PDF file.
#
#         Returns:
#             str: Status message indicating success or an error.
#         """
#         max_rows_per_page = 50
#         num_pages = len(self.df) / max_rows_per_page
#         num_pages = ceil(num_pages)
#
#         relative_column_widths = [0.3, 0.15, 0.1, 0.075, 0.075, 0.3]
#
#         pdf_file_name = f"inventory-{self.date_now.strftime('%d-%m-%Y-%H-%M')}.pdf"
#         pdf_file_path = f"{pdf_file_location}/{pdf_file_name}"
#
#         with PdfPages(pdf_file_path) as pdf:
#             for page in range(num_pages):
#                 fig, ax = plt.subplots(figsize=(8.27, 11.69))
#                 ax.axis('off')
#
#                 if page == 0:
#                     fig.text(0.5, 0.97, 'Inventory List', fontsize=16, fontweight='bold', ha='center', va='top')
#                     fig.subplots_adjust(top=0.99)
#
#                 start_row = page * max_rows_per_page
#                 end_row = min(start_row + max_rows_per_page, len(self.df))
#                 df_chunk = self.df.iloc[start_row:end_row]
#
#                 table = ax.table(cellText=df_chunk.values, colLabels=self.df.columns, cellLoc='center', loc='center', colWidths=relative_column_widths)
#                 table.auto_set_font_size(False)
#                 table.set_fontsize(6)
#
#                 table.scale(1, 1.2)
#                 ax.set_position([0, 0, 1, 0.9])
#
#                 pdf.savefig(fig, bbox_inches='tight')
#                 plt.close()
#
#         message = f"List Printed in the selected location with name - {pdf_file_name}"
#         return message
#
#
# class MyApp:
#     """
#     Tkinter GUI application for inventory management.
#     """
#
#     def __init__(self, root):
#         """
#         Initializes the Tkinter GUI.
#
#         Args:
#             root (tk.Tk): Tkinter root window.
#         """
#         self.root = root
#         self.root.title("Inventory Management")
#         self.root.geometry("400x350")
#         self.root.iconbitmap('Icon.ico')
#
#         # List to store items
#         self.item_list = []
#
#         # Set minimum window size
#         root.minsize(400, 350)
#
#         # Set maximum window size
#         root.maxsize(400, 350)
#
#         # Create and set up the GUI components
#         self.create_widgets()
#
#     def create_widgets(self):
#         """
#         Creates and sets up GUI components.
#         """
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#         tk.Button(self.root, text="Add", command=lambda: self.open_window("Add Item", ['Part Name', 'Part No', 'Model', 'Stock Location', 'Quantity'], self.add_item), padx=45, pady=2, font=("System", 8), width=6).pack()
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#         tk.Button(self.root, text="Remove", command=lambda: self.open_window("Remove Item", ['Part No', 'Quantity'], self.remove_item), padx=45, pady=2, font=("System", 8), width=6).pack()
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#         tk.Button(self.root, text="Search", command=lambda: self.open_window("Bring List", ['Part No'], self.search_item), padx=45, pady=2, font=("System", 8), width=6).pack()
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#         tk.Button(self.root, text="Bring List", command=self.bring_list, padx=45, pady=2, font=("System", 8), width=6).pack()
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#         tk.Button(self.root, text="Print List", command=self.print_list, padx=45, pady=2, font=("System", 8), width=6).pack()
#         tk.Label(self.root, text="", font=("System", 2)).pack()
#
#     def open_window(self, title, fields, command_function):
#         """
#         Opens a generic Tkinter window.
#
#         Args:
#             title (str): Title of the window.
#             fields (list): List of input fields.
#             command_function: Function to be executed when the "Submit" button is pressed.
#         """
#         generic_window = tk.Toplevel(self.root)
#         generic_window.title(title)
#         generic_window.geometry("400x350")
#         generic_window.iconbitmap('Icon.ico')
#         self.setup_input_fields(generic_window, fields, command_function)
#
#     def setup_input_fields(self, window, fields, command_function):
#         """
#         Sets up input fields in a Tkinter window.
#
#         Args:
#             window (tk.Toplevel): Tkinter window.
#             fields (list): List of input fields.
#             command_function: Function to be executed when the "Submit" button is pressed.
#         """
#         for field in fields:
#             tk.Label(window, text=f"{field}:").pack()
#             entry_var = tk.StringVar()
#             tk.Entry(window, textvariable=entry_var, name=f'{field.lower()}_entry').pack()
#
#         tk.Label(window, text="", font=("System", 2)).pack()
#         tk.Button(window, text="Submit", command=lambda: command_function(window, fields)).pack()
#
#     def add_item(self, window, fields):
#         """
#         Adds an item to the inventory.
#
#         Args:
#             window (tk.Toplevel): Tkinter window.
#             fields (list): List of input fields.
#         """
#         try:
#             part_name = str(window.children['part name_entry'].get())
#             part_no = str(window.children['part no_entry'].get())
#             model = str(window.children['model_entry'].get())
#             stock_location = int(window.children['stock location_entry'].get())
#             quantity = int(window.children['quantity_entry'].get())
#
#             if all((part_name, part_no, stock_location, quantity)):
#                 data = ["ADD", part_name, part_no, model, stock_location, quantity]
#                 self.item_list = data
#                 messagebox.showinfo("Success", "Item added successfully!")
#                 window.destroy()
#                 self.root.destroy()
#             else:
#                 messagebox.showwarning("Error", "Please fill in all fields.")
#
#         except ValueError:
#             messagebox.showerror("Invalid data type. Please enter valid data types for each field.")
#
#     def remove_item(self, window, fields):
#         """
#         Removes an item from the inventory.
#
#         Args:
#             window (tk.Toplevel): Tkinter window.
#             fields (list): List of input fields.
#         """
#         data = []
#         invalid_fields = []
#
#         for field in fields:
#             try:
#                 if field.lower() == 'part name' or field.lower() == 'part no' or field.lower() == 'model':
#                     field_input = str(window.children[field.lower() + '_entry'].get())
#                     data.append(field_input)
#                 else:
#                     field_input = int(window.children[field.lower() + '_entry'].get())
#                     data.append(field_input)
#             except ValueError:
#                 invalid_fields.append(field)
#
#         if invalid_fields:
#             self.show_error_message(f"Invalid data type. Please enter valid data types for {', '.join(invalid_fields)}.")
#         elif all(data):
#             data.insert(0, "REMOVE")
#             self.item_list = data
#             window.destroy()
#             self.root.destroy()
#
#     def search_item(self, window, fields):
#         """
#         Searches for an item in the inventory.
#
#         Args:
#             window (tk.Toplevel): Tkinter window.
#             fields (list): List of input fields.
#         """
#         data = [window.children[field.lower() + '_entry'].get() for field in fields]
#
#         if all(data):
#             data.insert(0, "SEARCH")
#             self.item_list = data
#             window.destroy()
#             self.root.destroy()
#         else:
#             messagebox.showwarning("Error", "Please fill in all fields.")
#
#     def bring_list(self):
#         """
#         Retrieves the entire inventory.
#
#         Returns:
#             list: Command to retrieve the inventory.
#         """
#         outlist = ["BRING"]
#         self.item_list = outlist
#         self.root.destroy()
#
#     def print_list(self):
#         """
#         Prints the inventory.
#
#         Returns:
#             list: Command to print the inventory.
#         """
#         folder_selected = filedialog.askdirectory()
#
#         if folder_selected:
#             outlist = ["PRINT", folder_selected]
#             self.item_list = outlist
#             self.root.destroy()
#         else:
#             self.show_error_message("Please select a folder")
#
#     def get_list(self):
#         """
#         Gets the list of items.
#
#         Returns:
#             list: List of items.
#         """
#         if self.item_list:
#             return self.item_list
#
#     def show_error_message(self, error_message):
#         """
#         Shows an error message.
#
#         Args:
#             error_message (str): Error message to display.
#         """
#         messagebox.showerror("Error", error_message)
#
#     def show_message(self, message):
#         """
#         Shows an information message.
#
#         Args:
#             message (str): Information message to display.
#         """
#         messagebox.showinfo("Message", message)
#
#     def display_excel_data(self, df, height):
#         """
#         Displays Excel data in a Tkinter window.
#
#         Args:
#             df (pd.DataFrame): DataFrame to display.
#             height (int): Height of the Tkinter window.
#         """
#         window = tk.Tk()
#         window.title("Data Display")
#         window.iconbitmap('Icon.ico')
#
#         text_widget = scrolledtext.ScrolledText(window, width=150, height=height, wrap=tk.WORD)
#         text_widget.pack(padx=10, pady=10)
#         text_widget.insert(tk.INSERT, df.to_string(index=False))
#         window.mainloop()

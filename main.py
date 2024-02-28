import tkinter as tk
from inventorymodule import Inventory as Inventory
from inventorymodule import MyApp as MyApp


# Create the main Tkinter root window
root = tk.Tk()

# Initialize the application
app = MyApp(root)

# Start the Tkinter event loop
root.mainloop()

# Get user input from the GUI
gui_input = app.get_list()

# Process user input and interact with the Inventory module
if gui_input:
    inventory = Inventory()

    action = gui_input[0]

    if action == "ADD":
        # Add item to the inventory
        return_str = inventory.add_method(*gui_input[1:])
        app.show_message(return_str[0], return_str[1])  # Display a success message
    elif action == "REMOVE":
        return_str = inventory.update_method(*gui_input[1:3])
        app.show_message(return_str[0], return_str[1])
    elif action == "BRING":
        excel_df = inventory.bring_list()
        app.display_excel_data(excel_df, 20)
    elif action == "SEARCH":
        result = inventory.search_method(gui_input[1])
        app.show_message(result[0], result[1]) if isinstance(result, str) else app.display_excel_data(result, 20)
    elif action == 'PRINT':
        return_str = inventory.print_list(gui_input[1])
        app.show_message(return_str[0], return_str[1])
    elif action == 'BULK':
        return_str = inventory.bulk_entry(gui_input[1])
        app.show_message(return_str[0], return_str[1])
    else:
        pass
else:
    pass
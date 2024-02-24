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
        message = inventory.add_method(*gui_input[1:])
        app.show_message("Message", message) # Display a success message
    elif action == "REMOVE":
        return_str = inventory.update_method(*gui_input[1:3])
        app.show_message("Message", return_str) if return_str == "Updated" else app.show_message("Error", return_str)
    elif action == "BRING":
        excel_df = inventory.bring_list()
        app.display_excel_data(excel_df, 20)
    elif action == "SEARCH":
        result = inventory.search_method(gui_input[1])
        app.show_message("Message", result) if isinstance(result, str) else app.display_excel_data(result, 20)
    elif action == 'PRINT':
        message = inventory.print_list(gui_input[1])
        app.show_message("Message", message)
    elif action == 'BULK':
        message = inventory.bulk_entry(gui_input[1])
        app.show_message("Message", message)
    else:
        app.show_message("Unkown Error Occured", "Message")













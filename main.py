import tkinter as tk
from inventorymodule import Inventory as Inventory
from inventorymodule import MyApp as MyApp


root = tk.Tk()
app = MyApp(root)
root.mainloop()

gui_input = app.get_list()

if gui_input:
    inventory = Inventory()

    action = gui_input[0]

    if action == "ADD":
        message = inventory.add_method(*gui_input[1:])
        app.show_message(message, "Message")
    elif action == "REMOVE":
        return_str = inventory.update_method(*gui_input[1:3])
        app.show_message(return_str, "Message") if return_str == "Updated" else app.show_message(return_str, "Error")
    elif action == "BRING":
        excel_df = inventory.bring_list()
        app.display_excel_data(excel_df, 20)
    elif action == "SEARCH":
        result = inventory.search_method(gui_input[1])
        app.show_message(result, "Message") if isinstance(result, str) else app.display_excel_data(result, 20)
    elif action == 'PRINT':
        message = inventory.print_list(gui_input[1])
        app.show_message(message, "Message")
    elif action == 'BULK':
        message = inventory.bulk_entry(gui_input[1])
        app.show_message(message, "Message")
    else:
        app.show_message("Unkown Error Occured", "Message")













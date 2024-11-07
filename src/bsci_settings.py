import json
import os
import tkinter as tk
import sys
import ctypes
import winsound

def show_windows_message(title, message):
    """Display a Windows message box with an OK button."""
    winsound.MessageBeep()
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)


app_data_folder = os.getenv('LOCALAPPDATA')
file_path = os.path.join(app_data_folder, "BMRAM_User_Job_Title\\BMRAM User, JobTitle_data\\settings.json")
print(file_path)

settings = {}
first_time = False
def first_time_file_was_loaded():
    data = {    
    "read_col": "A",
    "write_Email": "B",
    "write_Title": "C",
    "should_write_email": True,
    "should_write_title": True,
    "add_drop_down_filter": True,
    "open_when_done": False,
    "halt_on_error": False,
    "create_new_excel":False,
    "sheet": "Sheet1"
    }
    with open(file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)



if os.path.isfile(file_path): 
    pass
else:
    first_time_file_was_loaded()
    first_time = True


with open(file_path, 'r') as json_file:
        settings = json.load(json_file)

while len(settings) != 10:
    pass



def submit():
    # Handle the submission of the form
    
    if (emailBox.get().upper() == titleBox.get().upper() or 
        emailBox.get().upper() == readBox.get().upper() or 
        titleBox.get().upper() == readBox.get().upper()):
        show_windows_message("Error!", "columns must be different")
        return


    settings["should_write_email"] = checkbox1_var.get()
    settings["should_write_title"] = checkbox2_var.get()
    settings["add_drop_down_filter"] = checkbox3_var.get()
    settings["open_when_done"] = checkbox4_var.get()
    settings["halt_on_error"] = checkbox5_var.get()
    settings["create_new_excel"] = checkbox6_var.get()
    settings["read_col"] = readBox.get().upper().strip()
    settings["write_Email"] = emailBox.get().upper().strip()
    settings["write_Title"] = titleBox.get().upper().strip()
    settings["sheet"] = sheetBox.get().strip()
    if sheetBox.get().lower().strip() == "marcus is really cool":
        show_windows_message(":)", "Thanks!")
    elif sheetBox.get().lower() == "marcus is not really cool":
        show_windows_message(":(", ":(")
    with open(file_path, 'w') as json_file:
        json.dump(settings, json_file, indent=4)
    sys.exit(0)


root = tk.Tk()
root.title("SETTINGS")
root.geometry("300x300")  # Set the window size
root.resizable(False, False)
checkbox_vars = []



# Create each checkbox separately
checkbox1_var = tk.BooleanVar(value=settings["should_write_email"])
checkbox1 = tk.Checkbutton(root, text="Write Email", variable=checkbox1_var)
checkbox1.pack(anchor='w')
checkbox_vars.append(checkbox1_var)

checkbox2_var = tk.BooleanVar(value=settings["should_write_title"])
checkbox2 = tk.Checkbutton(root, text="Write JobTitle", variable=checkbox2_var)
checkbox2.pack(anchor='w')
checkbox_vars.append(checkbox2_var)

checkbox3_var = tk.BooleanVar(value=settings["add_drop_down_filter"])
checkbox3 = tk.Checkbutton(root, text="Add filter to Excel when done", variable=checkbox3_var)
checkbox3.pack(anchor='w')
checkbox_vars.append(checkbox3_var)

checkbox4_var = tk.BooleanVar(value=settings["open_when_done"])
checkbox4 = tk.Checkbutton(root, text="Open Excel when done", variable=checkbox4_var)
checkbox4.pack(anchor='w')
checkbox_vars.append(checkbox4_var)

checkbox5_var = tk.BooleanVar(value=settings["halt_on_error"])
checkbox5 = tk.Checkbutton(root, text="Halt on Error", variable=checkbox5_var)
checkbox5.pack(anchor='w')
checkbox_vars.append(checkbox5_var)

checkbox6_var = tk.BooleanVar(value=settings["create_new_excel"])
checkbox6 = tk.Checkbutton(root, text="Write to a new Excel file", variable=checkbox6_var)
checkbox6.pack(anchor='w')
checkbox_vars.append(checkbox6_var)




readBox = tk.Entry(root)
readBox.pack(pady=5)
readBox.place(x=100, y=155)
readBox.insert(0, settings["read_col"].upper())
emailBox = tk.Entry(root)
emailBox.pack(pady=5)
emailBox.place(x=100, y=185)
emailBox.insert(0, settings["write_Email"].upper())
titleBox = tk.Entry(root)
titleBox.pack(pady=5)
titleBox.place(x=100, y=215)
titleBox.insert(0, settings["write_Title"].upper())
sheetBox = tk.Entry(root)
sheetBox.pack(pady=5)
sheetBox.place(x=100, y=245)
sheetBox.insert(0, settings["sheet"])


label = tk.Label(root, text="Names column:")
label.place(x=10, y=155)
label = tk.Label(root, text="Email column:")
label.place(x=18, y=185)
label = tk.Label(root, text="Title column:")
label.place(x=25, y=215)
label = tk.Label(root, text="Sheet Name:")
label.place(x=27, y=245)


# Create a list for text boxes for easy access
text_boxes = [readBox, emailBox, titleBox, sheetBox]

# Create a submit button
submit_button = tk.Button(root, text="Save and Close", command=submit)
submit_button.place(x=105, y=275)


# Start the Tkinter event loop
root.mainloop()

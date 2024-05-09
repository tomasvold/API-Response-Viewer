import tkinter as tk
from tkinter import ttk
import requests
import json
from tkinter import filedialog
from tkinter import messagebox
import configparser
import pandas as pd
from ttkthemes import ThemedStyle

def load_last_credentials():
    config = configparser.ConfigParser()
    config.read('credentials.ini')
    client_id = config.get('Credentials', 'client_id', fallback='')
    refresh_token = config.get('Credentials', 'refresh_token', fallback='')
    return client_id, refresh_token

def exchange_refresh_token(event=None):
    client_id = client_id_entry.get()
    refresh_token = refresh_token_entry.get()

    url = 'https://id.planday.com/connect/token'
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'refresh_token': refresh_token
    }
    response = requests.post(url, headers=headers, data=data)
    response_data = response.json()

    access_token = response_data.get('access_token')
    bearer_entry.delete(0, tk.END)
    bearer_entry.insert(tk.END, access_token)

    # Save the entered Client ID and Refresh Token
    config = configparser.ConfigParser()
    config['Credentials'] = {'client_id': client_id, 'refresh_token': refresh_token}
    with open('credentials.ini', 'w') as configfile:
        config.write(configfile)

def make_api_call(event=None):
    url = url_dropdown.get()
    bearer_token = bearer_entry.get()
    headers = {
        'Authorization': f'Bearer {bearer_token}'
    }
    response = requests.get(url, headers=headers)
    try:
        data = response.json()
    except json.decoder.JSONDecodeError:
        text.delete('1.0', tk.END)
        text.insert(tk.END, "Error: Invalid JSON response")
        return

    # Clear previous data
    text.delete('1.0', tk.END)

    # Insert new data
    text.insert(tk.END, json.dumps(data, indent=4))

def download_response():
    response_text = text.get("1.0", tk.END)
    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if file_path:
        with open(file_path, "w") as file:
            file.write(response_text)

def find_text(event=None):
    find_dialog = tk.Toplevel(root)
    find_dialog.title("Find")
    find_label = ttk.Label(find_dialog, text="Find:")
    find_label.pack(side=tk.LEFT)
    find_entry = ttk.Entry(find_dialog, width=30)
    find_entry.pack(side=tk.LEFT, padx=10)
    find_entry.focus_set()

    find_results_text = tk.Text(find_dialog, wrap=tk.WORD, height=10)
    find_results_text.pack(fill=tk.BOTH, expand=True)

    def find():
        text_to_find = find_entry.get()
        start = "1.0"
        find_results_text.delete('1.0', tk.END)  # Clear previous results
        count = 0
        while True:
            start = text.search(text_to_find, start, stopindex=tk.END)
            if not start:
                break
            count += 1
            line_start = text.index(f"{start.split('.')[0]}.0")
            line_end = text.index(f"{start.split('.')[0]}.end")
            line_text = text.get(line_start, line_end)
            find_results_text.insert(tk.END, f"{line_text}\n")

            end = f"{start}+{len(text_to_find)}c"
            text.tag_add("found", start, end)
            start = end

        if count == 0:
            find_results_text.insert(tk.END, "No matches found")
        else:
            find_results_text.insert(tk.END, f"Number of results found: {count}")

        text.tag_config("found", background="yellow")

    find_button = ttk.Button(find_dialog, text="Find", command=find)
    find_button.pack()

    def close_find_dialog():
        text.tag_remove("found", "1.0", tk.END)
        find_dialog.destroy()

    find_dialog.protocol("WM_DELETE_WINDOW", close_find_dialog)

# Create the main window
root = tk.Tk()
root.title("Toms API Response Viewer")

# Apply a themed style
style = ThemedStyle(root)
style.set_theme("winnative")  # Choose a theme (e.g., 'equilux', 'arc')

# Set window size
window_width = 1000
window_height = 900
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_coordinate = (screen_width / 2) - (window_width / 2)
y_coordinate = (screen_height / 2) - (window_height / 2)
root.geometry("%dx%d+%d+%d" % (window_width, window_height, x_coordinate, y_coordinate))

# URL dropdown
url_frame = ttk.Frame(root)
url_frame.pack(pady=10)
url_label = ttk.Label(url_frame, text="SELECT ENDPOINT")
url_label.pack(side=tk.LEFT)

def show_tooltip(event):
    messagebox.showinfo("Information", "Query parameters can be added by appending the URL with ? for example /?createdFrom=2024-02-01&offset=50\n\nPath parameters can be added by appending the URL with /Id for example /1135")

# Information label
info_label = ttk.Label(url_frame, text="more info", foreground="blue", cursor="hand2")
info_label.pack(side=tk.LEFT, padx=1)
info_label.bind("<Button-1>", show_tooltip)

url_options = [
    "Please select",
    "https://openapi.planday.com/hr/v1.0/employees/",
    "https://openapi.planday.com/hr/v1.0/employees/deactivated",
    "https://openapi.planday.com/hr/v1.0/employees/fielddefinitions",
    "https://openapi.planday.com/hr/v1.0/departments/",
    "https://openapi.planday.com/hr/v1.0/employees/supervisors",
    "https://openapi.planday.com/hr/v1.0/employeegroups",
    "https://openapi.planday.com/hr/v1.0/employeetypes",
    "https://openapi.planday.com/hr/v1.0/employees/{employeeId}/history",
    "https://openapi.planday.com/hr/v1.0/skills",
    "https://openapi.planday.com/pay/v1.0/payrates/employeeGroups/{employeeGroupId}/employees/{employeeId}",
    "https://openapi.planday.com/pay/v1.0/payrates/employeeGroups/{employeeGroupId}/employees/{employeeId}/history",
    "https://openapi.planday.com/pay/v1.0/payrates/employees/{employeeId}",
    "https://openapi.planday.com/pay/v1.0/payrates/employeeGroups/default",
    "https://openapi.planday.com/pay/v1.0/salaries/employees/{employeeId}",
    "https://openapi.planday.com/pay/v1.0/salaries/types",
    "https://openapi.planday.com/punchclock/v1.0/employeeshifts/today",
    "https://openapi.planday.com/punchclock/v1.0/punchclockshifts/?From=yyyy-mm-ddT00:00&To=yyyy-mm-ddT00:00",
    "https://openapi.planday.com/absence/v1.0/absencetypes",
    "https://openapi.planday.com/absence/v1.0/accounts",
    "https://openapi.planday.com/absence/v1.0/accounts/{accountId}/adjustments",
    "https://openapi.planday.com/absence/v1.0/accounttypes",
    "https://openapi.planday.com/absence/v1.0/accounttypes/accrued/{id}",
    "https://openapi.planday.com/absence/v1.0/absencerecords",
    "https://openapi.planday.com/absence/v1.0/absencerequests",
    "https://openapi.planday.com/absence/v1.0/accounts/{id}/balance",
    "https://openapi.planday.com/absence/v1.0/accounts/{accountId}/transactions",
    "https://openapi.planday.com/portal/v1.0/info",
    "https://openapi.planday.com/revenue/v1.0/revenueunits",
    "https://openapi.planday.com/revenue/v1.0/departments/{departmentId}/revenueunits",
    "https://openapi.planday.com/revenue/v1.0/revenue",
    "https://openapi.planday.com/scheduling/v1.0/shifts",
    "https://openapi.planday.com/scheduling/v1.0/shifts/deleted",
    "https://openapi.planday.com/scheduling/v1.0/shifts/shiftstatus/all",
    "https://openapi.planday.com/scheduling/v1.0/shifts/{shiftId}/history",
    "https://openapi.planday.com/scheduling/v1.0/sections",
    "https://openapi.planday.com/scheduling/v1.0/positions",
    "https://openapi.planday.com/scheduling/v1.0/shifttypes",
    "https://openapi.planday.com/scheduling/v1.0/timeandcost/{departmentId}?from=yyyy-mm-dd&to=yyyy-mm-dd",
    "https://openapi.planday.com/pay/v1.0/salaries/types",
    "https://openapi.planday.com/pay/v1.0/salaries/scheduling/timeandcost/allocations/{employeeId}",
    "https://openapi.planday.com/scheduling/v1.0/scheduleDay?departmentId={departmentId}",
    "https://openapi.planday.com/payroll/v1.0/payroll/?departmentIds=xxxx&from=yyyy-mm-dd&to=yyyy-mm-dd"
]
url_dropdown = ttk.Combobox(url_frame, values=url_options, width=100)
url_dropdown.pack(side=tk.LEFT, padx=10)
url_dropdown.set(url_options[0])  # Set the default value

""" 
Endpoints excluded:
GET primary account
GET Absence request draft
GET Returns all contract rules present on a portal.
GET The attachment value of a custom field (image)
GET Return salary allocation history for a given employee id
GET Return a timeline list of Period Salaries with a given employee id
GET List of all Breaks on a Punch Clock record
GET SchedulingHistory
GET Security groups for portal 

QUERY PARAMS:
You can add Query parameters by appending the URL with ? and = in the URL, add multiple ones with &
for example: https://openapi.planday.com/hr/v1.0/employees?createdFrom=2024-02-02&offset=2
https://openapi.planday.com/hr/v1.0/employees/?departmentId=1135

PATH PARAMETERS:
You can add Path parameters by appending the URL with / followed by the ID.
for example: https://openapi.planday.com/hr/v1.0/departments/1136
"""

# Client ID input
client_id_frame = ttk.Frame(root)
client_id_frame.pack(pady=10)
client_id_label = ttk.Label(client_id_frame, text="INSERT CLIENT ID:")
client_id_label.pack(side=tk.LEFT)
client_id_entry = ttk.Entry(client_id_frame, width=50)
client_id_entry.pack(side=tk.LEFT, padx=10)

# Refresh Token input
refresh_token_frame = ttk.Frame(root)
refresh_token_frame.pack(pady=10)
refresh_token_label = ttk.Label(refresh_token_frame, text="INSERT REFRESH TOKEN:")
refresh_token_label.pack(side=tk.LEFT)
refresh_token_entry = ttk.Entry(refresh_token_frame, width=44)
refresh_token_entry.pack(side=tk.LEFT, padx=10)

# Load last entered Client ID and Refresh Token
last_client_id, last_refresh_token = load_last_credentials()
client_id_entry.insert(tk.END, last_client_id)
refresh_token_entry.insert(tk.END, last_refresh_token)

# Bearer token input
bearer_frame = ttk.Frame(root)
bearer_frame.pack(pady=10)
bearer_label = ttk.Label(bearer_frame, text="BEARER TOKEN:")
bearer_label.pack(side=tk.LEFT)
bearer_entry = ttk.Entry(bearer_frame, width=52)
bearer_entry.pack(side=tk.LEFT, padx=10)

# Create a button to trigger the Refresh Token exchange
exchange_button = tk.Button(root, text="Exchange Refresh Token", command=exchange_refresh_token)
exchange_button.pack()

# Create a button to trigger the API call
button = tk.Button(root, text="Make API Call", command=make_api_call)
button.pack()

# Create a Frame to hold the Text widget and Scrollbar
text_frame = tk.Frame(root)
text_frame.pack(expand=True, fill=tk.BOTH)

# Create a Text widget to display the JSON response
text = tk.Text(text_frame, wrap=tk.WORD)
text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

# Create a Scrollbar for the Text widget
scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
text.configure(yscrollcommand=scrollbar.set)

# Bind Ctrl+F to find_text function
text.bind("<Control-f>", find_text)

# Bind Return key to make_api_call function
root.bind("<Return>", make_api_call)

# Create a button to download the response as Json
download_button = tk.Button(root, text="Download Json response", command=download_response)
download_button.pack(side=tk.LEFT)

# Create a button to download and convert the response as Excel
def convert_to_excel():
    json_response = text.get("1.0", tk.END)
    try:
        data = json.loads(json_response)["data"]
        df = pd.DataFrame(data)

        excel_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if excel_path:
            df.to_excel(excel_path, index=False)
            messagebox.showinfo("Conversion Successful", "JSON response converted to Excel successfully!")
    except json.decoder.JSONDecodeError as e:
        messagebox.showerror("Conversion Error", "Invalid JSON format. Cannot convert to Excel.")

excel_button = tk.Button(root, text="Convert to Excel", command=convert_to_excel)
excel_button.pack(side=tk.RIGHT)

root.mainloop()

#to compile: python -m PyInstaller --onedir -w main.py
#use Inno setup compiler program to create the installation wizard program, include the main executable and the main folder inside of "Dist" directory
#last updated May 8 2024
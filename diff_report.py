from util_functions import getSlideNumbers, apiRunner, pd, generate_table, tk, openpyxl, threading

'''
Performs difference report on Excel sheet compared to API data
'''
def diffReport(root, text_field, button):
    file_path = 'APIData.xlsx' # Change to path where file is located

    dress_ids = sorted(getSlideNumbers(text_field)) # gets dress IDs in entry field
    diff_dress_data = [] # data in spreadsheet that is different from API
    api_dress_data = sorted(apiRunner(root, text_field), key=lambda x : x['id']) # gets dress data from API and sorts by ID

    try:
        # gets and cleans dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)
        sheet_dress_data.dropna(subset=['id'], inplace=True) # drops any rows with na ID
        sheet_dress_data['description'].fillna('', inplace=True) # removes na/nan from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True) # removes na/nan from did_you_know column
        sheet_dress_data['description'] = sheet_dress_data['description'].astype(str).apply(openpyxl.utils.escape.unescape) # convert escaped strings to ASCII
        sheet_dress_data['did_you_know'] = sheet_dress_data['did_you_know'].astype(str).apply(openpyxl.utils.escape.unescape) # Convert escaped strings to ASCII

        # cycle through api_dress_data
        for api_data in api_dress_data:
            # row of data with an ID that matches api_data ID
            row = sheet_dress_data.loc[api_data['id']-1] # row of data in spreadsheet

            # check if name, description, or did_you_know is different from the API data
            if row.loc['name'] != api_data['name']:
                new_row = [item for item in row]
                new_row.append('changed')
                diff_dress_data.append(new_row)
                continue
            if row.loc['description'] != api_data['description']:
                new_row = [item for item in row]
                new_row.append('changed')
                diff_dress_data.append(new_row)
                continue
            if row.loc['did_you_know'] != api_data['did_you_know']:
                new_row = [item for item in row]
                new_row.append('changed')
                diff_dress_data.append(new_row)
                continue

        # check for new entries in Excel sheet that do not exist in retrieved api data
        for id in dress_ids:
            if not any(data['id'] == id for data in api_dress_data) and (sheet_dress_data['id']==id).any():
                row = sheet_dress_data.loc[sheet_dress_data['id']==id]
                new_row = [item for item in row.values[0]]
                new_row.append('new')
                diff_dress_data.append(new_row)

        column_headers = ['id', 'name', 'description', 'did_you_know', 'changed_or_new']
        generate_table(root, diff_dress_data, 'difference_report', column_headers, 150, 800, 'nw')

    except FileNotFoundError:
        tk.messagebox.showerror(title="Error in diffReport", message=f"File '{file_path}' not found.")
        print(f"File '{file_path}' not found.")
    except Exception as e:
        tk.messagebox.showerror(title="Error in diffReport", message=f'Error: {e}')
        print(f'Error: {e}')
    finally:
        button.config(state="normal")

'''
Spins up new thread to run diffReport function
'''
def startDiffReportThread(root, text_field, button):
    button.config(state='disabled')
    diff_report_thread = threading.Thread(target=diffReport(root, text_field, button))
    diff_report_thread.start()
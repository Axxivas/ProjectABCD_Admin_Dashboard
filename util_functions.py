import webbrowser,threading, requests, os, urllib, platform, textwrap, time, openpyxl, string
import googletrans
import tkinter as tk
from tkinter import ttk
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import pandas as pd
from pandas import Series, DataFrame
import openai
from concurrent.futures import ThreadPoolExecutor, as_completed

'''
Gathers data from API
'''
def downloadAPIData(url, id_number):
    try:
        response = requests.get(url, headers={"User-Agent": "XY"})
        # append dress info to dress data if response status_code == 200
        if response.ok:
            return response.json()['data']
        else:
            print(f'Request for dress ID: {id_number} failed.')
    except requests.exceptions.RequestException as e:
        print(f'-- DEBUG -- in downloadAPIData: {e}')
    except Exception as e:
        # tk.messagebox.showerror(title="Error", message=f'Could not make connection!\n\nError: {e}')
        print(f'Error: {e}')

'''
Sets up and starts threads for gathering API data
'''
def apiRunner(root, text_field):
    # dress ids in entry field
    dress_ids = getSlideNumbers(text_field)

    # create list of all urls to send requests to
    url_list = []
    for id_number in dress_ids:
        url_list.append(f'https://abcd2.projectabcd.com/api/getinfo.php?id={id_number}')

    dress_data = [] # dress data from API
    threads= [] # working threads

    # create progress bar
    progress_window, pb, percent_label = progress_bar(root,'Retrieving API Data')

    # spins up 10 threads at a time and stores retrieved data into dress_data upon completion
    with ThreadPoolExecutor(max_workers=10) as exec:
        for index, url in enumerate(url_list):
            threads.append(exec.submit(downloadAPIData, url, dress_ids[index]))
        
        complete = 0 # number of threads that have finished
        for task in as_completed(threads):
            complete += 1
            pb['value'] = (complete/len(dress_ids))*100 # calculate percentage of data retrieved
            percent_label.config(text=f'Retrieving API Data...{int(pb["value"])}%') # update completion percent label

            if task.result() is not None:
                dress_data.append(task.result()) # append retrieved data to dress_data

    progress_window.destroy()
    return dress_data

'''
Gets dress IDs from entry field
'''
def getSlideNumbers(text_field):
    # get update for dress number input
    update_dress_list = []
    # get dress numbers from text field
    get_text_field = text_field.get("1.0", "end-1c").split(',')
    
    # add to list
    for number in get_text_field:
        if (number.strip().isnumeric()):
            update_dress_list.append(int(number.strip()))
    
    # remove duplicates
    dress_ids = []
    [dress_ids.append(x) for x in update_dress_list if x not in dress_ids]

    return dress_ids


'''
Downloads dress images
'''
def downloadImages(folder_name, url, img_name):
    try:
        # downloads dress image
        img_url = url
        img_path = f'./{folder_name}/{img_name}'
        opener = urllib.request.build_opener()
        opener.addheaders=[('User-Agent', 'XY')]
        urllib.request.install_opener(opener)
        urllib.request.urlretrieve(img_url, img_path)
    except Exception as e:
        print(f'Error downloading images: {e}')

'''
Sets up and starts threads for downloading dress images
'''
def imageRunner(root, dress_data):
    # list of all urls to send requests to
    url_list = []
    # list of all image names
    img_name_list = []

    for data in dress_data:
        url_list.append(f'http://projectabcd.com/images/dress_images/{data["image_url"]}')
        img_name_list.append(f'{data["image_url"]}')

    # create progress bar
    progress_window, pb, percent_label = progress_bar(root, 'Downloading Images')

    threads= [] # working threads

    # spins up 10 threads at a time and calls downloadImages with url and image name
    with ThreadPoolExecutor(max_workers=10) as exec:
        for index, url in enumerate(url_list):
            threads.append(exec.submit(downloadImages, "images", url, img_name_list[index]))

        complete = 0 # number of threads that have finished
        for task in as_completed(threads):
            complete += 1
            pb['value'] = (complete/len(url_list))*100 # calculate percentage of images downloaded
            percent_label.config(text=f'Downloading Images...{int(pb["value"])}%') # update completion percent label
    
    progress_window.destroy() # close progress bar window

'''
Generates progress bar
'''
def progress_bar(root, title):
    # progress bar window for API data retrieval
    progress_window = tk.Toplevel(root)
    progress_window.title(title)
    sw = int(progress_window.winfo_screenwidth()/2 - 450/2)
    sh = int(progress_window.winfo_screenheight()/2 - 70/2)
    progress_window.geometry(f'450x70+{sw}+{sh}')
    progress_window.resizable(False, False)
    progress_window.attributes('-disable', True)
    progress_window.focus()

    # progress bar custom style
    pb_style = ttk.Style()
    pb_style.theme_use('clam')
    pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

    # frame to hold progress bar
    pb_frame = tk.Frame(progress_window)
    pb_frame.pack()

    # progress bar
    pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
    pb.pack(pady=10)

    # label for percent complete
    percent_label = tk.Label(pb_frame, text=f'{title}...0%')
    percent_label.pack()

    return progress_window, pb, percent_label

'''
Opens file depending on OS
'''
def openFile(file_name):
    current_os = platform.system()
    try:
        if current_os == "Windows":
            print(f"-- DEBUG -- Windows {file_name}")
            os.system(f"start {file_name}")
        elif current_os == "Darwin":
            print(f"-- DEBUG -- Darwin {file_name}")
            os.system(f"open {file_name}")
        elif current_os == "Linux":
            print(f"-- DEBUG -- Linux {file_name}")
            os.system(f"xdg-open {file_name}")
        else:
            print("Error: Cannot open file " + current_os + " not supported.")
    except Exception as e:
        print("Error:", e)

'''
Closes the pop alert message
'''
def close_popup(popup):
    popup.destroy()

'''
Updates the timer for the alert popup
'''
def update_timer(root, popup, timer_label, seconds_left):
    timer_label.config(text=f"Auto closes in {seconds_left} sec.")

    if seconds_left > 0:
        # Update the timer every second
        root.after(1000, update_timer, popup, timer_label, seconds_left - 1)
    else:
        close_popup(popup)

'''
Displays the alert message
'''
def show_error_popup(root, text_message, duration_num):
    popup = tk.Toplevel(root)
    popup.title("Alert Message")
    
    label = tk.Label(popup, text=text_message)
    label.pack(padx=10, pady=10)

    timer_label = tk.Label(popup, text="")
    timer_label.pack(pady=5)

    duration = duration_num
    update_timer(popup, timer_label, duration)

'''
Helper function to wrap text
'''
def wrap(string, length=150):
    return '\n'.join(textwrap.wrap(string, length))

'''
Generates treeview table of data
'''
def generate_table(root,table_data, report_name, column_headers, row_h, col_w, anchor_point, num_buttons=2):
    # create window to display table
    table_window = tk.Toplevel(root)
    table_window.title(report_name)
    table_window.geometry(f"1000x600")
    table_window.minsize(1000,600)

    # create frame to hold table
    table_frame = tk.Frame(table_window)
    table_frame.pack_propagate(False)
    table_frame.place(x=0, y=0, relwidth=1, relheight=.89, anchor="nw")

    # using style to set row height and heading colors
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Treeview', rowheight=row_h)
    style.configure('Treeview.Heading', background='#848484', foreground='white')

    # vertical scrollbar
    table_scrolly = tk.Scrollbar(table_frame)
    table_scrolly.pack(side="right", fill='y')
    # horizontal scrollbar
    table_scrollx = tk.Scrollbar(table_frame, orient='horizontal')
    table_scrollx.pack(side="bottom", fill='x')

    # use ttk Treeview to create table
    table = ttk.Treeview(table_frame, yscrollcommand=table_scrolly.set, xscrollcommand=table_scrollx.set, columns=column_headers, show='headings')

    # configure the scroll bars with the table
    table_scrolly.config(command=table.yview)
    table_scrollx.config(command=table.xview)

    # create the headers and set column variables
    for index, column_header in enumerate(column_headers):
        table.heading(column_header, text=column_header)
        if index == 0:
            table.column(column_header, width=75, stretch=False)
        elif index == 1:
            table.column(column_header, width=145, stretch=False)
        else:
            table.column(column_header, width=col_w, stretch=False, anchor=anchor_point)

    # pack table into table_frame
    table.pack(fill='both', expand=True)

    # fill table with difference report data
    for index, data in enumerate(table_data):
        # word wrap text
        for i, cell in enumerate(data): 
            if len(str(data[i])) > 2000:
                data[i] = wrap(str(cell), 400)
            else:
                data[i] = wrap(str(cell), 250)

        # if new row, set tag to new
        # if changed row, set tag to changed
        if data[-1] == 'new':
            table.insert(parent='', index=tk.END, values=data, tags=('new',))
        elif data[-1] == 'changed':
            table.insert(parent='', index=tk.END, values=data, tags=('changed',))
        else:
            # if even row, set tag to evenrow
            # if odd row, set tag to oddrow
            if index % 2 == 0:
                table.insert(parent='', index=tk.END, values=data, tags=('evenrow',))
            else:
                table.insert(parent='', index=tk.END, values=data, tags=('oddrow',))

    # color rows
    table.tag_configure('new', background='#BAFFA4')
    table.tag_configure('changed', background='#FFA5A4')
    table.tag_configure('evenrow', background='#e8f3ff')
    table.tag_configure('oddrow', background='#f7f7f7')

    # create button frame and place one table_window
    btn_frame = tk.Frame(table_window)
    btn_frame.pack(side='bottom', pady=15)
    
    if num_buttons == 2:
        # create buttons and pack on button_frame
        btn = tk.Button(btn_frame, text='Export SQL File', font="helvetica bold", width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportSQL(table_data, column_headers, report_name))
        btn.pack(side='left', padx=25)

        btn2 = tk.Button(btn_frame, text='Export to HTML', font="helvetica bold", width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportHTML(table_data, column_headers, report_name))
        btn2.pack(side='left', padx=25)

    elif num_buttons == 1:
        btn = tk.Button(btn_frame, text='Export to HTML', font="helvetica bold", width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportHTML(table_data, column_headers, report_name))
        btn.pack(side='left', padx=25)

'''
Exports table data to Excel file
'''
def exportExcel(data, excel_columns, sheet_name):
    df = pd.DataFrame(data, columns=excel_columns)
    df.to_excel(f'{sheet_name}.xlsx', index=False)

'''
Exports difference report data to SQL update script
'''
def exportSQL(dress_data, column_headers, report_name):
    if report_name == 'difference_report':
        sql_queries = [] # stores sql query

        # cycle through the diff_dress_data and create queries for each
        for index, data in enumerate(dress_data):
            data[1] = str(data[1]).replace('"', '\\"') # replace " with \" so quotations don't mess up query
            data[2] = str(data[2]).replace('"', '\\"') # replace " with \" so quotations don't mess up query
            data[3] = str(data[3]).replace('"', '\\"') # replace " with \" so quotations don't mess up query
            if data[-1] == 'changed':
                sql_queries.append(f'UPDATE dresses\nSET name="{data[1]}", description="{data[2]}", did_you_know="{data[3]}"\nWHERE id={data[0]};\n')
            elif data[-1] == 'new':
                sql_queries.append(f'INSERT INTO dresses (id, name, description, did_you_know)\nVALUES ({data[0]}, "{data[1]}", "{data[2]}", "{data[3]}");\n')
        
        # create path for sql script
        update_script_path = 'abcdbook_SQL_update.sql'
        update_script_name = 'abcdbook_SQL_update'
        count = 1
        while os.path.exists(update_script_path):
            update_script_path = f'{update_script_name}({count}).sql'
            count += 1

        # write sql queries into .sql script
        with open(update_script_path, 'w') as f:
            for query in sql_queries:
                f.write(f'{query}\n')

    elif report_name == 'wiki_link_report':
        with open(f'{report_name}_update.sql', 'w') as sql_file:
            # Create SQL CREATE TABLE statement
            create_table_query = f'CREATE TABLE IF NOT EXISTS resources (\n'
            create_table_query += ', '.join(f'{column} TEXT' for column in column_headers)
            create_table_query += '\n);\n\n'
            sql_file.write(create_table_query)

            # Create SQL INSERT INTO statement
            sql_file.write(f'INSERT INTO resources ({", ".join(column_headers)}) VALUES\n')

            # Iterate through data and write values
            for row in dress_data:
                values = ', '.join(f"'{str(value)}'" for value in row)
                sql_file.write(f'({values}),\n')

            # Remove the trailing comma from the last line
            sql_file.seek(sql_file.tell() - 2)
            sql_file.truncate()

            # Add a semicolon to the end of the SQL script
            sql_file.write(';')

'''
Exports data to JQuery data table HTML page
'''
def exportHTML(data, column_headers, file_name):
    #read in data and create column headers
    table_data = pd.DataFrame(data)
    table_data.columns = column_headers

    #convert table_data to html
    html_table_data = table_data.to_html(table_id='html_table_data', border=0, classes='display')

    #html template to generate JQuery DataTable of data
    html_temp = f"""
    <!DOCTYPE html>
    <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=width-device, initial-scale=1.0">
            <title>{file_name}</title>
            <!--jQuery cdn-->
            <script src="https://code.jquery.com/jquery-3.7.0.js"></script>
            <!--Datatable style-->
            <link rel="stylesheet" href="https:////cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css" />
            <!--Datatable cdn-->
            <script src="https:////cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
            <!--Datatable button libraries-->
            <script src="https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js"></script>
            <script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/pdfmake.min.js"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.1.53/vfs_fonts.js"></script>
            <!--Initialize datatables-->
            <script>
                $(document).ready(function() {{
                    $('#html_table_data').DataTable({{
                        pageLength: 25,
                        dom: 'Bfrtip',
                        buttons: [
                            'csv', 'excel', 'pdf'
                        ]
                    }},
                    style_table = {{
                        'width': '100%'
                    }});
                }});
            </script>
        </head>
        <body>
            {html_table_data}
        </body>
    </html>
    """

    #write html template into datatable.html file
    with open(f'{file_name}.html', 'w', encoding='utf-8') as f:
        f.write(html_temp)

    #open datatable.html
    webbrowser.open(f'{file_name}.html')

'''
Sorts dress data based on user selection
'''
def sortDresses(sort_order, dress_data):
    if sort_order.get() == 1:
        return sorted(dress_data, key=lambda x : str(x['name']).lower())
    elif sort_order.get() == 2:
        return sorted(dress_data, key=lambda x : x['id'])
    elif sort_order.get() == 3:
        return dress_data

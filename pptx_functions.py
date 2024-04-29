'''
There are 3 buttons in this file
Generate Book
Generate First Person Text
Generate Word Search Puzzle
'''


'''
Performs update once generate button clicked
'''
def generateBook():
    # check if Generate from Local is active
    if gen_local.get() == 1:
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

        # path to local Excel data
        file_path = "APIData.xlsx"

        # gets and cleans dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)
        sheet_dress_data.dropna(subset=['id'], inplace=True) # drops any rows with na ID
        sheet_dress_data['description'].fillna('', inplace=True) # removes na/nan from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True) # removes na/nan from did_you_know column
        sheet_dress_data['description'] = sheet_dress_data['description'].astype(str).apply(openpyxl.utils.escape.unescape) # convert escaped strings to ASCII
        sheet_dress_data['did_you_know'] = sheet_dress_data['did_you_know'].astype(str).apply(openpyxl.utils.escape.unescape) # Convert escaped strings to ASCII
        
        # holds dress data from local Excel sheet
        dress_data = [] 
        # cycle through dress_ids and append data from excel sheet to dress_data
        for id in dress_ids:
            row = sheet_dress_data.loc[id-1]
            dress_data.append({'id':row.loc['id'], 'name':row.loc['name'], 'description':row.loc['description'], 'did_you_know':row.loc['did_you_know']})   
    else:
        dress_data = apiRunner() # gather all dress data from api
    
    # sort dress data
    sorted_dress_data = sortDresses(dress_data)

    # if there is no image in the local folder then download from web
    if download_imgs.get() == 0:
        if not os.path.exists('./images'):
            os.makedirs('./images')
        if not os.listdir('./images'):
            text_message = "Local folder is empty. Attempting to grab image(s) from web."
            duration_num = 4
            show_error_popup(text_message, duration_num)
            time.sleep(duration_num+1)
            imageRunner(sorted_dress_data)

    # download images from web if download images check box is selected
    if download_imgs.get() == 1:
        # creates directory to save images if one does not exist
        if not os.path.exists('./images'):
            os.makedirs('./images')
        imageRunner(sorted_dress_data) # download images for each dress in list

    # create powerpoint
    prs = Presentation() # create the pptx presentation
    ppt_file_name = "abcdbook.pptx"
    file_name = "abcdbook.pptx"
    count = 0
    while os.path.exists(file_name): # check if file name exist,
        count += 1
        file_name = f"{os.path.splitext(ppt_file_name)[0]}({count}).pptx" # if file name exisit create new filename

    # progress bar window for API data retrieval
    progress_window = tk.Toplevel(root)
    progress_window.title('Creating Book')
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
    percent_label = tk.Label(pb_frame, text='Creating Book...0%')
    percent_label.pack()
    complete = 0

    # get dress info for items in list & translate
    for index, dress_info in enumerate(sorted_dress_data):
        left = None
        image_left = None
        
        dress_name = dress_info['name']
        dress_description = dress_info['description']
        dress_did_you_know = dress_info['did_you_know']
        dress_description_len = len(dress_description)

        #--------------------------------Portrait--------------------------------
        # PORTRAIT MODE
        if layout.get() == 1 or layout.get() == 4: 
            prs.slide_width = pptx.util.Inches(7.5) # define slide width
            prs.slide_height = pptx.util.Inches(10.83) # define slide height
            slide_layout = prs.slide_layouts[5] # use slide with only title
            slide_layout2 = prs.slide_layouts[6] # use empty slide

            # LAYOUT 1 == picture on left page - text on right page - two page
            if layout.get() == 1:
                slide_empty = prs.slides.add_slide(slide_layout2) 
                slide_title = prs.slides.add_slide(slide_layout) 

                add_image(slide_empty, dress_info, 0, 0)
                add_title_box(slide_title, dress_name, 0, 0.15, 7.5, 0.91) 
                add_subtitle_highlight(slide_title, 0.37, 1.58, 2.44, 0.3) # description - highlight box
                add_description_subtitle(slide_title, 0.28, 1.07, 6.94, 0.51)
                add_description_text(slide_title, dress_description, 0.28, 1.65, 6.94, 5.99)
                add_subtitle_highlight(slide_title, 0.37, 8.36, 2.78, 0.3) # did you know - highlight box
                add_did_you_know_subtitle(slide_title, 0.28, 7.87, 6.94, 0.51)
                add_did_you_know_text(slide_title, dress_did_you_know, 0.28, 8.46, 6.94, 1.04)
                add_numbering(slide_title, dress_info, index, 4.47, 10.06, 1.28, 0.34, 5.94, 10.06, 1.28, 0.34)

            # LAYOUT 4 == picture on left - text on right - single page
            elif layout.get() == 4: 
                slide_title = prs.slides.add_slide(slide_layout) 

                add_image(slide_title, dress_info, 0, 1.39)
                add_title_box(slide_title, dress_name, 0, 0.09, 7.5, 0.91) 
                add_subtitle_highlight(slide_title, 3.71, 1.21, 1.83, 0.23) # description - highlight box
                add_description_subtitle(slide_title, 3.63, 0.88, 3.71, 0.37)
                add_description_text(slide_title, dress_description, 3.63, 1.35, 3.71, 7.57)

                # adjust the text height based on text length
                if dress_description_len < 600:
                    add_subtitle_highlight(slide_title, 3.72, 5.83, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 5.42, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 5.94, 3.71, 0.91)

                elif dress_description_len > 600 and dress_description_len < 1300:
                    add_subtitle_highlight(slide_title, 3.72, 7.99, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 7.58, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 8.1, 3.71, 0.91)
                    
                else:
                    add_subtitle_highlight(slide_title, 3.72, 9.32, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 8.91, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 9.43, 3.71, 0.91)

                add_numbering(slide_title, dress_info, index, 0.49, 6.46, 1.33, 0.27, 1.95, 6.46, 1.33, 0.27)
 
        #--------------------------------Landscape--------------------------------
        # LANDSCAPE MODE
        elif layout.get() == 2 or layout.get() == 3:
            left = None
            image_left = None
            rectangle_left = None

            # slide size (left, top, width, height)
            prs.slide_width = pptx.util.Inches(16.18) # define slide width
            prs.slide_height = pptx.util.Inches(12.53) # define slide height
            slide_layout = prs.slide_layouts[5] # use empty slide
            slide = prs.slides.add_slide(slide_layout) # add empty slide to pptx

            # LAYOUT 2 == picture on right - text on left
            if layout.get() == 2:
                rectangle_left = 0.4
                left = 0.34
                image_left = 8.45
                image_top = 1.17
                numbering1_left = 1.05
                numbering2_left = 2.53

            # LAYOUT 3 == picture on left - text on right
            elif layout.get() == 3:
                image_left = 0.25
                image_top = 1.17
                rectangle_left = 8.15
                left = 8.09
                numbering1_left = 12.78
                numbering2_left = 14.26

            add_title_box(slide, dress_name, 0, 0.15, 16.18, 0.91) 
            add_subtitle_highlight(slide, rectangle_left, 1.75, 2.44, 0.3) # decription - highlight box
            add_description_subtitle(slide, left, 1.26, 7.81, 0.51)
            add_description_text(slide, dress_description, left, 1.88, 7.81, 5.35)
            add_subtitle_highlight(slide, rectangle_left, 8.04, 2.76, 0.28) # did you know - highlight box
            add_did_you_know_subtitle(slide, left, 7.56, 7.81, 0.51)
            add_did_you_know_text(slide, dress_did_you_know, left, 8.19, 7.81, 1.11)
            add_numbering(slide, dress_info, index, numbering1_left, 11.08, 1.28, 0.34, numbering2_left, 11.08, 1.28, 0.34)
            add_image(slide, dress_info, image_left, image_top)
        complete += 1
        pb['value'] = (complete/len(sorted_dress_data))*100 # calculate percentage of images downloaded
        percent_label.config(text=f'Creating Book...{int(pb["value"])}%') # update completion percent label
    try:
        prs.save(file_name)
    except Exception as e:
        print(f"-- DEBUG -- saving presentation: {e}")
    finally:
        book_gen_generate_button.config(state="normal")
    
    progress_window.destroy() # close progress bar window

    openFile(file_name)
    print(sorted_dress_data)


def read_translated_ids(filename):
    translated_ids = set()
    if os.path.exists(filename):
        with open(filename, 'r') as file:
            for line in file:
                translated_ids.add(line.strip())
    return translated_ids

def write_translated_ids(translated_ids, filename):
    with open(filename, 'w') as file:
        for id in translated_ids:
            file.write(f"{id}\n")

'''
Performs change of third person text to first person
'''
def translate_text_to_first_person(english_texts):
    i = 1
    tries = 1
    sleep_timer = 0
    first_person_texts = {} # Dictionary to hold the translated text
    translated_ids_file = 'translated_ids.txt'
    translated_ids = read_translated_ids(translated_ids_file)
    print(translated_ids, flush=True)


    with open("output.txt", "a") as file: # Opens a text file to write the translated text
        pass

    messages = [{"role": "system", "content": 
                 "You translate text from third person to first person."}]


    for id, texts in english_texts.items():
        if str(id) in translated_ids:
            print(f"ID: {id} was found to be translated already.\n\n", flush=True)
            continue

        first_person_texts[id] = {}
        for key, text in texts.items():
            if key == "description" or key == "did_you_know":
                try:

                    # Perform conversion from third person to first person using ChatGPT
                    messages.append(
                            {"role": "user", "content": f"Convert the following text to first person:\n\n{text}\n\n"}
                            )
                    chat = openai.ChatCompletion.create(
                        model="gpt-3.5-turbo",
                        messages=messages
                    )
                    reply = chat.choices[0].message.content
                    first_person_texts[id][key] = reply

                    print(f"The ID: {id}\nThe KEY: {key}\nThe TEXT: {text}\nThe REPLY: {reply}\nTries: {tries}\n\n", flush=True)
                    tries += 1
                    if tries > 40:
                        break
                except Exception as e:
                    sleep_timer += 1
                    time.sleep(sleep_timer)
                    print("\n\nError Occurred: \n", e, flush=True)

                    #first_person_texts[id] = first_person_text.choices[0].text.strip()
            else:
                first_person_texts[id][key] = text

        if 'description' in first_person_texts[id] and 'did_you_know' in first_person_texts[id]:
            translated_ids.add(id)
            write_translated_ids(translated_ids, translated_ids_file)
            with open("output.txt", "a") as file:
                file.write(f"Original text from ID: {id}\n")
                file.write(f"Description:\n{english_texts[id]['description']}\n")
                file.write(f"Did You Know:\n{english_texts[id]['did_you_know']}\n\n")
                file.write(f"First-person text from ID: {id}\n")
                file.write(f"Description:\n{first_person_texts[id]['description']}\n")
                file.write(f"Did You Know:\n{first_person_texts[id]['did_you_know']}\n")
                file.write("======================================================================================\n")
            print(f"======================\nCharacter {i} Translated\n======================\n", flush=True)
            i+= 1
            print(f"Sleeper Timer at: {sleep_timer}", flush=True)
            print("+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++", flush=True)
        if tries > 40:
            break
    print("\n\nDONE WITH EVERYTHING\n\n", flush=True)
    return first_person_texts

'''
Create powerpoint for first person text
'''
def first_person_pptx(first_person_text):
    try:
        dress_data = apiRunner() # gather all dress data from api
    
        # sort dress data
        sorted_dress_data = sortDresses(dress_data)

        # if there is no image in the local folder then download from web
        if download_imgs.get() == 0:
            if not os.path.exists('./images'):
                os.makedirs('./images')
            if not os.listdir('./images'):
                text_message = "Local folder is empty. Attempting to grab image(s) from web."
                duration_num = 4
                show_error_popup(text_message, duration_num)
                time.sleep(duration_num+1)
                imageRunner(sorted_dress_data)

        # download images from web if download images check box is selected
            # creates directory to save images if one does not exist
        if not os.path.exists('./images'):
            os.makedirs('./images')
        imageRunner(sorted_dress_data) # download images for each dress in list
        prs = Presentation()
        ppt_file_name = "fpp.pptx"
        file_name = "fpp.pptx"
        count = 0
        while os.path.exists(file_name):
            count += 1
            file_name = f"{os.path.splitext(ppt_file_name)[0]}({count}).pptx"

        progress_window = tk.Toplevel(root)
        progress_window.title('Creating First Person Pptx')
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
        percent_label = tk.Label(pb_frame, text='Creating Powerpoint...0%')
        percent_label.pack()
        complete = 0

        for id, dress_info in first_person_text.items():
            dress_data = {'id': id}
            print("ID", id, flush=True)
            print("DRESS INFO: ", dress_info, flush=True)
            dress_name = first_person_text[id].get('name', '')
            dress_description = first_person_text[id].get('description', '')
            dress_did_you_know = first_person_text[id].get('did_you_know', '')

            dress_description_len = len(dress_description)



            #--------------------------------Portrait--------------------------------
            prs.slide_width = pptx.util.Inches(7.5) # define slide width
            prs.slide_height = pptx.util.Inches(10.83) # define slide height
            slide_layout = prs.slide_layouts[5] # use slide with only title
            slide_layout2 = prs.slide_layouts[6] # use empty slide

            slide_title = prs.slides.add_slide(slide_layout) 

            add_image(slide_title, dress_data, 0, 1.39)
            add_title_box(slide_title, dress_name, 0, 0.09, 7.5, 0.91) 
            add_subtitle_highlight(slide_title, 3.71, 1.21, 1.83, 0.23) # description - highlight box
            add_description_subtitle(slide_title, 3.63, 0.88, 3.71, 0.37)
            add_description_text(slide_title, dress_description, 3.63, 1.35, 3.71, 7.57)

                # adjust the text height based on text length
            if dress_description_len < 600:
                add_subtitle_highlight(slide_title, 3.72, 5.83, 1.83, 0.23) # did you know - highlight box
                add_did_you_know_subtitle(slide_title, 3.63, 5.42, 3.71, 0.37)
                add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 5.94, 3.71, 0.91)

            elif dress_description_len > 600 and dress_description_len < 1300:
                add_subtitle_highlight(slide_title, 3.72, 7.99, 1.83, 0.23) # did you know - highlight box
                add_did_you_know_subtitle(slide_title, 3.63, 7.58, 3.71, 0.37)
                add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 8.1, 3.71, 0.91)
                
            else:
                add_subtitle_highlight(slide_title, 3.72, 9.32, 1.83, 0.23) # did you know - highlight box
                add_did_you_know_subtitle(slide_title, 3.63, 8.91, 3.71, 0.37)
                add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 9.43, 3.71, 0.91)

            #add_numbering(slide_title, dress_info, id, 0.49, 6.46, 1.33, 0.27, 1.95, 6.46, 1.33, 0.27)
            complete += 1
            #pb['value'] = (complete/len(first_person_text))*100 # calculate percentage of images downloaded
            #percent_label.config(text=f'Creating Book...{int(pb["value"])}%')
            dress_data.clear()
        prs.save(file_name)
        openFile(file_name)
        progress_window.destroy()
    except Exception as e:
        traceback.print_exc()

def create_html_package_gpt(english_texts, first_person_texts):
    page = 1
    html_content = """
    <html>
    <head>
    <style>
        .name { font-weight: bold; text-align: center; }
        .did_you_know { margin-top: 20px; font-size: 18px; }
        .page {
            page-break-after: always;
        }
        @media screen {
            .page {
                border-bottom: 1px solid #ccc;
                padding-bottom: 20px;
                margin-bottom: 20px;
            }
        }
    </style>
    </head>
    <body>
    """
    for id, english_text in english_texts.items():
        first_person_text = first_person_texts.get(id, {'name': '', 'description': '', 'did_you_know': ''})
        html_content += f"<div class='page'><h2>Page No: {page} ABCD-id: {id} (English)</h2>"
        html_content += f"<div class='name'>{english_text['name']}</div>"
        html_content += f"<div class='description>'>Description</div>"
        html_content += f"<p>{english_text['description']}</p>"
        html_content += "<div class='did_you_know'>Did you know?</div>"
        html_content += f"<p>{english_text['did_you_know']}</p></div><hr>"
        
        html_content += f"<div class='page'><h2>Page No: {page} ABCD-id: {id} (ChatGPT)</h2>"
        html_content += f"<div class='name'>{english_text['name']}</div>"
        html_content += f"<div class='description>'>Description</div>"
        html_content += f"<p>{first_person_text['description']}</p>"
        html_content += "<div class='did_you_know'>Did you know?</div>"
        html_content += f"<p>{first_person_text['did_you_know']}</p></div><hr>"
        page += 1
    html_content += "</body></html>"
    return html_content

def save_html_to_file_gpt(english_texts, first_person_texts, base_filename="gpt_adjusted_package"):
    html_filename = f"{base_filename}.html"
    txt_filename = f"{base_filename}.txt"
    counter = 1
    # Check if the file exists and update the filename until it's unique
    while os.path.exists(html_filename) or os.path.exists(txt_filename):
        html_filename = f"{base_filename}_{counter}.html"
        txt_filename = f"{base_filename}_{counter}.txt"
        counter += 1
    
    # Saving HTML content
    with open(html_filename, 'w', encoding='utf-8') as file:
        file.write(create_html_package_gpt(english_texts, first_person_texts))  # Assuming create_html_package is defined as before
    
    # Creating a plain text version of the content
    text_content = ""
    for id, english_text in english_texts.items():
        telugu_text = first_person_texts.get(id, '')
        text_content += f"English Text for ID {id}\n{english_text}\n\n"
        text_content += "------------------------------------------------\n\n"
        text_content += f"ChatGPT Text for ID {id}\n{telugu_text}\n\n"
        text_content += "================================================\n\n"
    
    # Saving text content
    with open(txt_filename, 'w', encoding='utf-8') as file:
        file.write(text_content)

    return html_filename, txt_filename

def generate_first_person_package():
    try:

        dress_data = apiRunner() 
        english_texts = fetch_english_text(dress_data)
        first_person_texts = translate_text_to_first_person(english_texts)
        first_person_pptx(first_person_texts)
        create_html_package_gpt(english_texts, first_person_texts)
        html_filename, txt_filename = save_html_to_file_gpt(english_texts, first_person_texts)
        webbrowser.open(f'file://{os.path.realpath(html_filename)}')
    except FileNotFoundError:
        print(f"File '{e}' not found")
    except Exception as e:
        print(f'Error: {e}')
    finally:
        first_person_generate_button.config(state='normal')
    print("Translation package has been generated and saved.", flush=True)
    
def wordSearchOpenAi(english_texts):
    words_for_puzzles = {}
    word_count = int(word_count_var.get())  # Get the current preferred word count

    messages = [
        {"role": "system", "content": f"You will extract {word_count} meaningful words significant to the character's text. Do not reply with anything else. Just list the words."}
    ]

    #sorted_texts = sort_english_texts(english_texts)  # Sort texts based on preference

    for id, text in english_texts.items():
        messages.append(
            {"role": "user", "content": f"Can you extract {word_count} words from this character's text that are meaningful and significant to the character, with one being their name? Please only output the {word_count} words.\n\n{text}\n\n"}
        )
        
        chat_response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=messages
        )
       
        reply = chat_response.choices[0].message.content.strip()
        
        words_for_puzzles[id] = reply
        print(f"This is the Reply for ID {id}: {reply}\n\n")
    
    return words_for_puzzles


#function that splits the words and ids, preparing them for puzzle creation
def wordsearchCreator(words_for_puzzles):
    puzzles = {}
    answer_keys = {}  
    word_lists = {}
    for id, words_string in words_for_puzzles.items():
        word_list = words_string.replace(',', '').replace("'", "").upper().split()
        word_lists[id] = word_list
        grid_size = int(puz_width_var.get())  # Get grid size
        grid = [['-' for _ in range(grid_size)] for _ in range(grid_size)]
        answer_positions = {}

        for word in word_list:
            placed, positions = placeWord(grid, word)
            if placed:
                # Store all positions, not just start and end
                answer_positions[word] = positions

        fillEmptySpots(grid)
        puzzles[id] = grid
        answer_keys[id] = answer_positions
    
    return puzzles, answer_keys, word_lists



#function to randomly place words in the grid
def placeWord(grid, word):
    max_attempts = 100    
    attempts = 0
    placed = False
    positions = []  

    while not placed and attempts < max_attempts:
        wordPlacement = random.randint(0, 3)
        attempts += 1

        if wordPlacement == 0:  # Horizontal
            row = random.randint(0, len(grid) - 1)
            col = random.randint(0, len(grid) - len(word))
            space_available = all(grid[row][col + i] == '-' or grid[row][col + i] == word[i] for i in range(len(word)))
            if space_available:
                for i in range(len(word)):
                    grid[row][col + i] = word[i]
                    positions.append((row, col + i))  
                placed = True

        elif wordPlacement == 1:  # Vertical
            row = random.randint(0, len(grid) - len(word))
            col = random.randint(0, len(grid) - 1)
            space_available = all(grid[row + i][col] == '-' or grid[row + i][col] == word[i] for i in range(len(word)))
            if space_available:
                for i in range(len(word)):
                    grid[row + i][col] = word[i]
                    positions.append((row + i, col))  
                placed = True

        elif wordPlacement == 2:  # Diagonal left to right
            row = random.randint(0, len(grid) - len(word))
            col = random.randint(0, len(grid) - len(word))
            space_available = all(grid[row + i][col + i] == '-' or grid[row + i][col + i] == word[i] for i in range(len(word)))
            if space_available:
                for i in range(len(word)):
                    grid[row + i][col + i] = word[i]
                    positions.append((row + i, col + i)) 
                placed = True

        elif wordPlacement == 3:  # Diagonal right to left
            row = random.randint(0, len(grid) - len(word))
            col = random.randint(len(word) - 1, len(grid) - 1)
            space_available = all(grid[row + i][col - i] == '-' or grid[row + i][col - i] == word[i] for i in range(len(word)))
            if space_available:
                for i in range(len(word)):
                    grid[row + i][col - i] = word[i]
                    positions.append((row + i, col - i))  
                placed = True

    return placed, positions

#function to fill the empty slots after words are placed                 
def fillEmptySpots(grid):
    for row in range(len(grid)):
        for col in range(len(grid[0])):  
            if grid[row][col] == '-':  
                grid[row][col] = random.choice(string.ascii_uppercase)
            
def createWordsearchWordsHtml(puzzles, answer_keys, word_lists):
    page = 1
    html_content = """
    <html>
    <head>
    <title>Word Search Puzzles</title>
    <style>
        body {
            font-family: Arial, sans-serif;
        }
        .page {
            page-break-after: always;
            margin-bottom: 20px;
        }
        table {
            border-collapse: collapse;
            margin: 20px 0;
            position: relative;
            border: 1px solid red; /* Temporary for alignment checking */
        }
        td {
            border: 1px solid #666;
            width: 20px;
            height: 20px;
            text-align: center;
            vertical-align: middle;
            box-sizing: border-box;
            color: black; /* Default text color */
        }
        .answer-text {
            color: blue; /* Change the text color to blue */
            font-weight: bold; /* Make answer text bold */
        }
    </style>
    </head>
    <body>
    """

    # Regular puzzles and word lists
    for id, grid in puzzles.items():
        html_content += f"<div class='page'><h2>Page No: {page} - Puzzle ID: {id}</h2><table>"
        for row in grid:
            html_content += "<tr>"
            for cell in row:
                html_content += f"<td>{cell}</td>"
            html_content += "</tr>"
        html_content += "</table><div><strong>Words:</strong><ul>"
        for word in word_lists[id]:
            html_content += f"<li>{word}</li>"
        html_content += "</ul></div></div>"
        page += 1

    # Answer key puzzles with text color change for answers only
    for id, positions in answer_keys.items():
        grid = puzzles[id]
        html_content += f"<div class='page'><h2>Answer Key Page No: {page} - Puzzle ID: {id}</h2><table>"
        flat_positions = set(sum(positions.values(), []))  # Flatten list of tuples from all words into a set for quick lookup
        for row_idx, row in enumerate(grid):
            html_content += "<tr>"
            for col_idx, cell in enumerate(row):
                # Apply the 'answer-text' class if the cell position is in the flat list of answer positions
                if (row_idx, col_idx) in flat_positions:
                    html_content += f"<td class='answer-text'>{cell}</td>"
                else:
                    html_content += f"<td>{cell}</td>"
            html_content += "</tr>"
        html_content += "</table></div>"
        page += 1

    html_content += "</body></html>"
    return html_content






def save_and_display_html(html_content, base_filename="puzzles_package"):
    html_filename = f"{base_filename}.html"
    counter = 1
    
    # Increment filename if exists to avoid overwriting
    while os.path.exists(html_filename):
        html_filename = f"{base_filename}_{counter}.html"
        counter += 1

    # Save HTML to file
    with open(html_filename, 'w') as file:
        file.write(html_content)
    print(f"HTML content has been saved to {html_filename}.")
    
    # Format the file path for browser compatibility and open it
    try:
        file_url = f"file://{os.path.abspath(html_filename)}"
        webbrowser.open(file_url, new=2)
        print("HTML file has been opened in your web browser.")
    except Exception as e:
        print(f"Failed to open the HTML file in a web browser. Error: {e}")
        
def add_puzzle_table(slide, grid, word_list, title_text, max_width, max_height, answer_positions=None):
    title = slide.shapes.title
    title.text = title_text

    rows, cols = len(grid), len(grid[0])
    grid_origin_x = Inches(1)
    grid_origin_y = Inches(1.5)

    # Calculate the cell width and height based on the max dimensions provided
    cell_width = max_width / cols
    cell_height = max_height / rows

    # Add the grid table
    table = slide.shapes.add_table(rows, cols, grid_origin_x, grid_origin_y, round(cell_width * cols), round(cell_height * rows)).table
    table.first_row = False

    for r in range(rows):
        for c in range(cols):
            cell = table.cell(r, c)
            cell.text = grid[r][c]
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(10)  
            p.alignment = PP_ALIGN.CENTER
            if answer_positions and any((r, c) in positions for positions in answer_positions.values()):
                p.font.color.rgb = RGBColor(255, 0, 0)  

    # Add lines for correct answers using direct line drawing
    if answer_positions:
        for positions in answer_positions.values():
            for i in range(len(positions)-1):
                start_cell = positions[i]
                end_cell = positions[i+1]
                start_x = grid_origin_x + start_cell[1] * cell_width + cell_width / 2
                start_y = grid_origin_y + start_cell[0] * cell_height + cell_height / 2
                end_x = grid_origin_x + end_cell[1] * cell_width + cell_width / 2
                end_y = grid_origin_y + end_cell[0] * cell_height + cell_height / 2

                
                line = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, start_x, start_y, end_x - start_x, end_y - start_y)
                line.line.width = Pt(2)
                line.line.color.rgb = RGBColor(255, 0, 0)  

    # Add word list to the side
    textbox = slide.shapes.add_textbox(Inches(7.5), Inches(1.5), Inches(2), Inches(4))
    tf = textbox.text_frame
    p = tf.add_paragraph()
    p.text = "Words to find:\n" + "\n".join(word_list)
    p.font.bold = True
    p.font.size = Pt(14)




def make_powerpoint(puzzles, answer_keys, word_lists):
    prs = Presentation()
    max_width = Inches(6)  # Set the maximum width for the grid
    max_height = Inches(4.5)  # Set the maximum height for the grid

    for id, grid in puzzles.items():
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        add_puzzle_table(slide, grid, word_lists[id], f"Puzzle ID: {id}", max_width, max_height)

    for id, positions in answer_keys.items():
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        add_puzzle_table(slide, puzzles[id], word_lists[id], f"Answer Key ID: {id}", max_width, max_height, answer_positions=positions)

    return prs






def save_powerpoint(prs, base_filename="puzzles_package"):
    filename = f"{base_filename}.pptx"
    counter = 1
    
    # Increment filename if exists to avoid overwriting
    while os.path.exists(filename):
        filename = f"{base_filename}_{counter}.pptx"
        counter += 1

    # Save PowerPoint to file
    prs.save(filename)
    print(f"PowerPoint content has been saved to {filename}.")
    
def generate_word_search_package():
    
    dress_data = apiRunner()

    
    english_texts = fetch_english_text(dress_data)

   
    puzzleWords = wordSearchOpenAi(english_texts)

    
    puzzles, answer_keys, word_lists = wordsearchCreator(puzzleWords)

    
    html_puzzles = createWordsearchWordsHtml(puzzles, answer_keys, word_lists)

   
    save_and_display_html(html_puzzles)
    print("HTML word puzzles have been generated and saved.")

   
    prs = make_powerpoint(puzzles, answer_keys, word_lists)

    
    save_powerpoint(prs, base_filename="word_puzzles_package")
    print("PowerPoint word puzzles have been generated and saved.")

'''
Spins up new thread to run generateBook
'''
def startGenerateBookThread():
    book_gen_generate_button.config(state="disabled")
    generate_thread = threading.Thread(target=generateBook)
    generate_thread.start()

'''
Spins up new thread to run translate_to_first person function
'''
def startFirstPersonThread():
    first_person_generate_button.config(state="disabled")
    first_person_thread = threading.Thread(target=generate_first_person_package)
    first_person_thread.start()

'''
Spins up new thread to run word search function
'''
def startWordPuzzleThread():
    word_puzzle_generate_button.config(state="disabled")
    word_search_thread = threading.Thread(target=generate_word_search_package)
    word_search_thread.start()

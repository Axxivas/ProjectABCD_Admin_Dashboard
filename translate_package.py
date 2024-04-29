from util_functions import googletrans, os, apiRunner, webbrowser, threading

'''
Function to fetch the english text from dress data
'''
def fetch_english_text(dress_data):
    english_texts = {}  
    for dress in dress_data:
        dress_id = dress.get('id')
        english_description = {
            'name': dress.get('name', ''),
            'description': dress.get('description', ''),
            'did_you_know': dress.get('did_you_know', '')
        }
        english_texts[dress_id] = english_description
    return english_texts
'''
translation function that stores the translated telugu texts into the same "keys" as the english for easier formatting
'''
def translate_text_to_telugu(english_texts):
    translator = googletrans.Translator()
    telugu_texts = {}

    for id, texts in english_texts.items():
        telugu_texts[id] = {}  # Initialize a dictionary for this ID
        for key, text in texts.items():
            try:
                translated = translator.translate(text, dest='te')  # 'te' for Telugu
                telugu_texts[id][key] = translated.text
            except Exception as e:
                print(f"Error translating {key} for ID {id}: {e}")
                telugu_texts[id][key] = text  # Use the original text if translation fails

    return telugu_texts

"""
    Create an HTML package with English and Telugu texts.
    """
def create_html_package(english_texts, telugu_texts):
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
        telugu_text = telugu_texts.get(id, {'name': '', 'description': '', 'did_you_know': ''})
        html_content += f"<div class='page'><h2>Page No: {page} ABCDid: {id} (English)</h2>"
        html_content += f"<div class='name'>{english_text['name']}</div>"
        html_content += f"<p class='description'>{english_text['description']}</p>"
        html_content += "<div class='did_you_know'>Did you know?</div>"
        html_content += f"<p>{english_text['did_you_know']}</p></div><hr>"
        
        html_content += f"<div class='page'><h2>Page No: {page} ABCDid: {id} (Telugu)</h2>"
        html_content += f"<div class='name'>{telugu_text['name']}</div>"
        html_content += f"<p class='description'>{telugu_text['description']}</p>"
        html_content += "<div class='did_you_know'>Did you know?</div>"
        html_content += f"<p>{telugu_text['did_you_know']}</p></div><hr>"
        page += 1
    html_content += "</body></html>"
    return html_content

"""
    Save the HTML content to a file, appending an incrementing number to the filename
    to avoid overwrites.
    """
def save_html_to_file(english_texts, telugu_texts, base_filename="translation_package"):
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
        file.write(create_html_package(english_texts, telugu_texts))  # Assuming create_html_package is defined as before
    
    # Creating a plain text version of the content
    text_content = ""
    for id, english_text in english_texts.items():
        telugu_text = telugu_texts.get(id, '')
        text_content += f"English Text for ID {id}\n{english_text}\n\n"
        text_content += "------------------------------------------------\n\n"
        text_content += f"Telugu Text for ID {id}\n{telugu_text}\n\n"
        text_content += "================================================\n\n"
    
    # Saving text content
    with open(txt_filename, 'w', encoding='utf-8') as file:
        file.write(text_content)

    return html_filename, txt_filename

def generate_translation_package(root, text_field, button):
    # Get ID numbers from text area
    try:

        dress_data = apiRunner(root, text_field)
        
        # Fetch English text
        english_texts = fetch_english_text(dress_data)
        

        # Translate to Telugu
        telugu_texts = translate_text_to_telugu(english_texts)
        
        # Create HTML package
        html_content = create_html_package(english_texts, telugu_texts)
    
        
        # Save HTML and TXT files
        html_filename, txt_filename = save_html_to_file(english_texts, telugu_texts)
    
        
        # Open the HTML file in a web browser
        webbrowser.open(f'file://{os.path.realpath(html_filename)}')
    
    except FileNotFoundError:
        print(f"File '{e}' not found.")
    except Exception as e:
        print(f'Error: {e}')
    finally:
        button.config(state='normal')
        
        # Optionally, show a message that the file has been saved
        print("Translation package has been generated and saved.")

'''
Spins up new thread to run translatepackage function
'''
def startTranslationPackageThread(root, text_field, button):
    button.config(state="disabled")
    translate_package_thread = threading.Thread(target=generate_translation_package(root, text_field, button))
    translate_package_thread.start()

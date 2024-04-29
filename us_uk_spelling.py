from localspelling import convert_spelling
from localspelling.spelling_converter import get_dictionary
from textblob import TextBlob
from util_functions import apiRunner, generate_table, pd, threading

def generate_us_uk_spellings(root, text_field, button):
    file_path = 'APIData.xlsx'  # Change to the actual file path
    api_dress_data = sorted(apiRunner(root, text_field), key=lambda x: x['id'])  # Assuming apiRunner() returns dress data
    words = []

    try:
        # Read dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)      
        sheet_dress_data.dropna(subset=['id'], inplace=True)  # Drop rows with missing IDs
        sheet_dress_data['description'].fillna('', inplace=True)  # Remove NA/NAN from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True)  # Remove NA/NAN from did_you_know column

	# Iterate through API dress data
        for api_data in api_dress_data:
            ID = api_data['id']
            name = api_data['name']
            description = api_data['description']
            did_you_know = api_data['did_you_know']

            # concatenation of description and did_you_know
            merged_words = f'{str(description)} {str(did_you_know)}'
            blob = TextBlob(merged_words)

            uk_words = []
            us_words = []
            uk_version_of_the_us_words = []
            same_words_for_both = []

            # Iterate over each word in the text
            for word in blob.words:
                # to check if it is us or uk
                if word in get_dictionary("us").values():
                    us_words.append(word)
                    uk_version_of_the_us_words.append(convert_spelling(word, 'gb'))

                elif word in get_dictionary("gb").values():
                    uk_words.append(word)
                else:
                    same_words_for_both.append(word)     

            words.append([ID, name, us_words, uk_version_of_the_us_words])

        # Generate table and save to Excel
        column_headers = ['ID', 'Name', 'US spellings', 'UK spellings']
        generate_table(root, words, 'US_UK Spellings', column_headers, 50, 300, 'center', 1)
        df_pairs = pd.DataFrame(words, columns=column_headers)
        df_pairs.to_excel("US_UK_Spelling.xlsx", index=False)
        print("Excel file 'US_UK_Spelling.xlsx' created.")

    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f'Error: {e}')
    button.config(state='normal')

# ids_of_us_uk spelling function
def generate_IDs_Of_us_uk_spellings(root, text_field, button):

    from collections import defaultdict

    file_path = 'APIData.xlsx'  # Change to the actual file path
    api_dress_data = sorted(apiRunner(root, text_field), key=lambda x: x['id'])  # Assuming apiRunner() returns dress data
    # Initialize a defaultdict to store dress IDs for each US spelling
    us_spellings_dict = defaultdict(lambda: {'ids': []})

    try:
        # Read dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)      
        sheet_dress_data.dropna(subset=['id'], inplace=True)  # Drop rows with missing IDs
        sheet_dress_data['description'].fillna('', inplace=True)  # Remove NA/NAN from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True)  # Remove NA/NAN from did_you_know column

        unique_ids_words_pairs = []  # Initialize an empty list to store the lists

        for api_data in api_dress_data:
            ID = api_data['id']
            description = api_data['description']
            did_you_know = api_data['did_you_know']

            # Concatenation of description and did_you_know
            merged_words = f'{str(description)} {str(did_you_know)}'
            blob = TextBlob(merged_words)

            uk_words = []
            us_words = []
            uk_version_of_the_us_words = []
            same_words_for_both = []

            # Iterate over each word in the text
            for word in blob.words:
                # to check if it is us or uk
                if word in get_dictionary("us").values():
                    us_words.append(word)
                    uk_version_of_the_us_words.append(convert_spelling(word, 'gb'))

                    # Append dress ID to the corresponding US spelling
                    us_spellings_dict[word]['ids'].append(ID)

                elif word in get_dictionary("gb").values():
                    uk_words.append(word)
                else:
                    same_words_for_both.append(word)

        # Create the list of unique_ids_words_pairs
        for us_word, data in us_spellings_dict.items():
            ids = data['ids']
            unique_ids = list(set(ids))  # Convert to set to remove duplicates, then back to list
            uk_word = convert_spelling(us_word, 'gb')
            
            # Create a new list with the US word, UK spelling, and unique IDs
            pair = [us_word, uk_word, unique_ids]
            
            # Append the list to the list
            unique_ids_words_pairs.append(pair)

        # Generate table and save to Excel
        column_headers = ['US', 'UK', 'IDS']
        generate_table(root, unique_ids_words_pairs, 'IDS_OF_US_UK Spellings', column_headers, 50, 300, 'center', 1)
        df_pairs = pd.DataFrame(unique_ids_words_pairs, columns=column_headers)
        df_pairs.to_excel("IDS_OF_US_UK Spellings.xlsx", index=False)
        print("Excel file 'IDS_OF_US_UK Spellings.xlsx' created.")

    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f'Error: {e}')
    button.config(state='normal')

'''
Spins up new thread to run us/uk_spellings function
'''
def startUS_UK_SpellingsThread(root, text_field, button):
    button.config(state="disabled")
    us_uk_spellings_thread = threading.Thread(target=generate_us_uk_spellings(root, text_field, button))
    us_uk_spellings_thread.start()

'''
Spins up new thread to run ids of us/uk_spellings function
'''
def startIDs_Of_US_UK_SpellingsThread(root, text_field, button):
    button.config(state="disabled")
    ids_of_us_uk_spellings_thread = threading.Thread(target=generate_IDs_Of_us_uk_spellings(root, text_field, button))
    ids_of_us_uk_spellings_thread.start()

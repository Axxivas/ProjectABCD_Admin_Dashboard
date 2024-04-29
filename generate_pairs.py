from util_functions import apiRunner, pd, generate_table, threading

def generatePairs(root, text_field, button):
    file_path = 'APIData.xlsx'  # Change to the actual file path
    api_dress_data = sorted(apiRunner(root, text_field), key=lambda x: x['id'])  # Assuming apiRunner() returns dress data
    pairs = []

    try:
        # Read dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)      
        sheet_dress_data.dropna(subset=['id'], inplace=True)  # Drop rows with missing IDs
        sheet_dress_data['description'].fillna('', inplace=True)  # Remove NA/NAN from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True)  # Remove NA/NAN from did_you_know column

	# Iterate through API dress data
        for data_that_will_be_searched in api_dress_data:
            # Get the name of the current API data ID
            name = data_that_will_be_searched['name']
            # Split the name into tokens and remove any prefixes
            tokens = [token + " " for token in name.split() if token not in ["Dr.", "Mr.", "Mrs.", "Ms."]]

            # Loop through each token
            for token in tokens:
                # Loop through provided IDs instead of the entire sheet_dress_data
                for data_that_contains_token in api_dress_data:
                    # Get the description and did you know text of the provided ID
                    description = data_that_contains_token['description']
                    did_you_know = data_that_contains_token['did_you_know']

                    # Check if the token is present in either of them, make sure we aren't looking on the same IDs
                    if data_that_will_be_searched['id'] != data_that_contains_token['id']:
                        if token in description or token in did_you_know:
                            if (data_that_contains_token['id'], data_that_will_be_searched['id']) not in [(pair[0], pair[2]) for pair in pairs]:
                                # Add the pair of IDs and names to the list
                                pairs.append([data_that_contains_token['id'], data_that_contains_token['name'], data_that_will_be_searched['id'], name])
                            # Break the inner loop as we found a pair for this token
                            break

        # Generate table and save to Excel
        column_headers = ['ID1', 'Name 1', 'ID2', 'Name 2']
        generate_table(root, pairs, 'generate_pairs', column_headers, 50, 200, 'center', 1)
        df_pairs = pd.DataFrame(pairs, columns=column_headers)
        df_pairs.to_excel("pairs_generated.xlsx", index=False)
        print("Excel file 'pairs_generated.xlsx' created.")

    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    except Exception as e:
        print(f'Error: {e}')
    button.config(state='normal')

'''
Spins up new thread to run generatePairs function
'''
def startGeneratePairsThread(root, text_field, button):
    button.config(state='disabled')
    who_are_my_pairs_thread = threading.Thread(target=generatePairs(root, text_field, button))
    who_are_my_pairs_thread.start()
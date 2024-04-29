import textstat
from textblob import TextBlob
from util_functions import apiRunner, tk, generate_table, string, threading

'''
Performs word analysis on given dress IDs
'''
def wordAnalysis(root, text_field, button):
    word_analysis_data = [] # data from word analysis
    api_dress_data = sorted(apiRunner(root, text_field), key=lambda x : x['id']) # gets dress data from API and sorts by ID

    try:
        # cycle through api_dress_data for word analysis
        for dress_data in api_dress_data:
            noun_count = 0 # number of nouns in text
            adjective_count = 0 # number of adjectives in text

            # concatenation of description and did_you_know
            text = f'{str(dress_data["description"])} {str(dress_data["did_you_know"])}'
            # TextBlob word analysis
            blob = TextBlob(text)

            # cycle through key:value pairs of TextBlob analysis to get noun and adjective count
            for k,v in blob.tags:
                if v == 'NN' or v == 'NNS' or v == 'NNP' or v == 'NNPS':
                    noun_count += 1
                elif v == 'JJ' or v == 'JJR' or v == 'JJS':
                    adjective_count += 1

            ease = textstat.flesch_reading_ease(text)
            kincaid = textstat.flesch_kincaid_grade(text)
            readability = textstat.automated_readability_index(text)

            # data to be displayed in table
            word_analysis_data.append([dress_data['id'], dress_data['name'], len(str(dress_data['description']).strip(string.punctuation).split()), 
                                       len(str(dress_data['did_you_know']).strip(string.punctuation).split()), str(noun_count), str(adjective_count),
                                       str(ease), str(kincaid), str(readability)])

        column_headers = ['id', 'name', 'description_word_count', 'did_you_know_word_count', 'total_noun_count', 'total_adjective_count', 'reading_ease', 'kincaid_grade', 'readability_index']
        generate_table(root, word_analysis_data, 'word_analysis_report', column_headers, 50, 200, 'center', 1)
        
    except Exception as e:
        tk.messagebox.showerror(title="Error in wordAnalysis", message=f'Error: {e}')
        print(f'Error: {e}')
    finally:
        button.config(state='normal')

'''
Spins up new thread to run wordAnalysis function
'''
def startWordAnalysisThread(root, text_field, button):
    button.config(state='disabled')
    word_analysis_thread = threading.Thread(target=wordAnalysis(root, text_field, button))
    word_analysis_thread.start()


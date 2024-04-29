import wikipediaapi
from util_functions import apiRunner, progress_bar, generate_table, threading

'''
Generate Wiki Link
'''
def generateWikiLink(root, text_field, button):
    # set user agent
    USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    wiki_wiki = wikipediaapi.Wikipedia("en",headers={"User-Agent": USER_AGENT})

    # gather all dress data from api
    wiki_link_data = sorted(apiRunner(root, text_field), key=lambda x : x['id'])
    wiki_data = []

    progress_window, pb, percent_label = progress_bar(root, 'Retrieving Wiki Data')

    for complete, item in enumerate(wiki_link_data):
        try:
            page = wiki_wiki.page(item["name"])
            if page.exists():
                item["wiki_page_link"] = page.fullurl
                wiki_data.append([item['id'], item['name'],item['wiki_page_link']])

                pb['value'] = (complete/len(wiki_link_data))*100 # calculate percentage of images downloaded
                percent_label.config(text=f'Retrieving Wiki Data...{int(pb["value"])}%') # update completion percent label
        except Exception as e:
            print(f"Error retrieving Wikipedia data for {item['name']}: {e}")

    progress_window.destroy()

    column_headers = ['id', 'name', 'wiki_page_link']
    generate_table(root, wiki_data, 'wiki_link_report', column_headers, 75, 800, 'nw')

    button.config(state='normal')

'''
Spins up new thread to run generateWikiLink function
'''
def startGenerateWikiLinkThread(root, text_field, button):
    button.config(state='disabled')
    wiki_link_thread = threading.Thread(target=generateWikiLink(root, text_field, button))
    wiki_link_thread.start()

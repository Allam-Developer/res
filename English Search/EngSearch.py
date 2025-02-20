from os import startfile
from re import search
from threading import Thread
from time import sleep
import tkinter as tk
from requests import session

# Global Variables
if True:
    search_type = 0
    all_or_common = 0
    current_page = 1
    current_letters = ""
    current_url = ""
    current_max_pages = 0
    columns = 2

    highlight_color = "red"

    file_text = ""

    current_session = session()

# highlight all letters in the scroll text
def highlight_all(word):
    global highlight_color, search_type

    if search_type == 3:
        for tag in scroll_text.tag_names():
            scroll_text.tag_remove(tag, "1.0", tk.END)

        for letter in word:
            start = "1.0"
            while True:
                pos = scroll_text.search(letter, start, tk.END)
                if not pos:
                    break
                end = f"{pos}+{1}c"
                scroll_text.tag_add(f"{letter}_tag", pos, end)
                start = end
            scroll_text.tag_config(f"{letter}_tag", foreground=highlight_color)
    else:
        for tag in scroll_text.tag_names():
            scroll_text.tag_remove(tag, "1.0", tk.END)
        start = "1.0"
        while True:
            pos = scroll_text.search(word, start, tk.END)
            if not pos:
                break
            end = f"{pos}+{len(word)}c"
            scroll_text.tag_add("highlight", pos, end)
            start = end
        scroll_text.tag_config("highlight", foreground=highlight_color)

# show the data on the ui
def add_page_to_ui(items, page):
    splitter = 1
    space = 30
    page_mark = "-"*20

    if not items: return

    scroll_text.insert(tk.END,f"\n{page_mark}( page {page} ){page_mark}\n\n")
    for item in items:
        item = item.lower()
        if splitter < columns:
            scroll_text.insert(tk.END, f"{item}{' '* (space-len(item))}")
            splitter += 1
        else:
            scroll_text.insert(tk.END, f"{item}\n")
            splitter = 1

# generate the text to be ready for creating the txt file
def generate_text():
    global file_text
    file_text = ''
    splitter = 1
    space = 30

    progress_label = tk.Label(results_frame, text=f"pages saved : 0/{current_max_pages}", font=("Helvetica", 10))
    progress_label.pack(side="top", pady=(0, 0))

    options = ["Contains", "Starts With", "Ends With", "Any Order"]
    options2 = ["All", "Common"]
    file_name = f"{options2[all_or_common]} words {options[search_type].lower()} {current_letters}.txt"


    for i in range(1, current_max_pages + 1):
        if i == 1:
            words = read_page(current_letters, i, search_type, all_or_common)[0]
        else:
            words = read_page(current_letters, i, search_type, all_or_common)

        if not words: return

        for word in words:
            word = word.lower()
            if splitter < columns:
                file_text = f"{file_text}{word}{' ' * (space - len(word))}"
                splitter += 1
            else:
                file_text = f"{file_text}{word}\n"
                splitter = 1

        progress_label.config(text=f"pages saved : {i}/{current_max_pages}")
    progress_label.config(text=f"done saved as\n {file_name}")
    with open(file_name, "w", encoding="utf-8") as file:
        file.write(file_text)
    startfile(file_name)
    sleep(5)
    progress_label.destroy()

# save all pages of the current search as a txt file
def save_as_txt():
    saving_file_thread = Thread(target=generate_text)
    saving_file_thread.start()

# every time a button released when the text field is in focus
def update_text(event=None):
    global scroll_text, search_type, current_letters, all_or_common
    text = text_entry.get()
    scroll_text.delete("1.0", tk.END)
    if text:
        current_letters = text
        page_data = read_page(text, 1, search_type, all_or_common)
        words = page_data[0]
        last_page = page_data[1]
        match_words_count = page_data[2]
        common_words_count = page_data[3]
        add_page_to_ui(words,1)
        highlight_all(text)
        results_label.config(text=f"words : {match_words_count}  ({common_words_count} common)      pages : {last_page}")

# make a http request based on the user inputs, get a html page as a response , extract the data from it :)
def read_page(letters, page, search_type=0,all_or_common=0): # search_type = contains, starts, ends
    global current_url, current_max_pages
    search_type_options = ["contains", "begins", "ends", "any-order"]
    all_or_common_options = ["all", "common"]

    search_type = search_type_options[search_type]
    all_or_common = all_or_common_options[all_or_common]
    url = f"https://www.merriam-webster.com/wordfinder/classic/{search_type}/{all_or_common}/-1/{letters}/{page}"
    current_url = url
    response = current_session.get(url)
    html_lines = response.text.splitlines()
    words = []
    first_word = 0
    last_word = 0
    for line in html_lines:
        if '<a href="/dictionary/' in line:
            if first_word == 0:
                first_word = html_lines.index(line)
            match = search(r'>([^<]+)<', line)
            words.append(match.group(1))
            last_word = html_lines.index(line)
    if page != 1:
        return words

    if first_word > 0 and last_word > 0:
        del html_lines[first_word:last_word]

    try:
        match = search(r'of\s+(\d+)', html_lines[546])
        last_page = int(match.group(1))
    except Exception:
        last_page = 1
    current_max_pages = last_page

    try:
        match = search(r'title="([\d,]+)"', html_lines[532])
        match2 = search(r'title="([\d,]+)"', html_lines[536])
        match_words = match.group(1)
        common_words = match2.group(1)
    except Exception:
        match_words = "0"
        common_words = "0"
        last_page = 0
    return words, last_page, match_words, common_words

# read the next page :)
def read_next_page():
    global current_page, current_letters, search_type
    text = text_entry.get()
    if text and current_page+1 <= current_max_pages:
        page_data = read_page(current_letters, current_page+1, search_type)
        current_page += 1
        add_page_to_ui(page_data, current_page)
        highlight_all(current_letters)

# open the link of the original website that data is taken from
def show_original_page():
    import webbrowser
    webbrowser.open(current_url)

# when search type dropdown is changed
def search_type_changed(value):
    global search_type
    options = ["Contain", "Start With", "End With", "Any Order"]
    search_type = options.index(value)
    update_text()

# when all or common dropdown is changed
def all_or_common_changed(value):
    global all_or_common
    options = ["All", "Common"]
    all_or_common = options.index(value)
    update_text()

# when columns dropdown is changed
def columns_changed(value):
    global columns
    options = ["1 Column", "2 Columns", "3 Columns", "4 Columns", "5 Columns"]
    columns = options.index(value) + 1
    update_text()

# load next page if you reached the end of the current page
def on_scroll(first, last):
    scrollbar.set(first, last)
    if float(last) >= 0.9999:
        read_next_page()

# UI
if True:
    root = tk.Tk()
    root.title("English Search")
    root.geometry("600x500")

    content_frame = tk.Frame(root)
    content_frame.pack(fill="both", expand=True, padx=10, pady=10)

    entry_frame = tk.Frame(content_frame)
    entry_frame.pack(fill="both", pady=(0, 40))
    text_entry = tk.Entry(entry_frame, width=50)
    text_entry.pack(side="left", padx=(0, 5))
    search_type_option = tk.StringVar()
    search_type_option.set("Contains")
    all_or_common_option = tk.StringVar()
    all_or_common_option.set("All")
    columns_option = tk.StringVar()
    columns_option.set("2 Columns")
    search_type_dropdown = tk.OptionMenu(entry_frame, search_type_option, "Contain", "Start With", "End With", "Any Order", command=search_type_changed)
    search_type_dropdown.pack(side="left", padx=(0, 0))
    all_or_common_dropdown = tk.OptionMenu(entry_frame, all_or_common_option, "All", "Common", command=all_or_common_changed)
    all_or_common_dropdown.pack(side="left", padx=(0, 0))
    columns_dropdown = tk.OptionMenu(entry_frame, columns_option, "1 Column", "2 Columns", "3 Columns", "4 Columns", "5 Columns", command=columns_changed)
    columns_dropdown.pack(side="left", padx=(0, 0))
    text_entry.bind('<KeyRelease>', update_text)

    results_frame = tk.Frame(content_frame)
    results_frame.pack(fill="both", pady=(0, 5))
    results_label = tk.Label(results_frame, text="results", font=("Helvetica", 10))
    results_label.pack(side="left", pady=(0, 0))
    original_page_button = tk.Button(results_frame,command=show_original_page, text="original page", font=("Helvetica", 8))
    original_page_button.pack(side="right", padx=(2, 2))
    txt_file_button = tk.Button(results_frame,command=save_as_txt, text="txt file", font=("Helvetica", 8))
    txt_file_button.pack(side="right", padx=(2, 2))

    scroll_text_frame = tk.Frame(content_frame)
    scroll_text_frame.pack(fill="both", expand=True)
    scrollbar = tk.Scrollbar(scroll_text_frame)
    scrollbar.pack(side="right", fill="y")
    scroll_text = tk.Text(scroll_text_frame,wrap="none", yscrollcommand=on_scroll, height=10)
    scroll_text.pack(side="left", fill="both", expand=True)
    scrollbar.config(command=scroll_text.yview)

    root.mainloop()

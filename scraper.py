import os
from tkinter import messagebox, ttk
import requests
from bs4 import BeautifulSoup
import re
import random
import pandas as pd
import logging
import warnings
from logging import config
import sys
import time
import tkinter as tk
from tkinter import filedialog
import openpyxl
warnings.filterwarnings("ignore")

config.dictConfig(
    {
        "version": 1,
        "disable_existing_loggers": False,
        "formatters": {
            "default": {
                "format": "[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] %(message)s"
            },
            "slack_format": {
                "format": "`[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d]` %(message)s"
            },
        },
        "handlers": {
            "file": {
                "class": "logging.FileHandler",
                "formatter": "default",
                "filename": "logs.log",
            },
        },
        "loggers": {
            "root": {
                "level": logging.INFO,
                "handlers": ["file"],
                "propagate": False,
            },
        },
    }
)
log = logging.getLogger("root")


def collect_query_list(df):
    queries = []
    for index, row in df.iterrows():
        name = row[0]
        unique_number = str(row[1]).zfill(8)
        queries.append(f"{name} {unique_number}")
    return queries


def get_proxy():
    proxies_list = []
    with open('proxies.txt', 'r') as file:
        for line in file:
            proxies_list.append(line.strip())
    ip, port, username, password = random.choice(proxies_list).split(":")
    proxy = {
        "http": f"http://{username}:{password}@{ip}:{port}",
        "https": f"https://{username}:{password}@{ip}:{port}",
    }
    log.info(f"Proxy in use: {proxy}")
    return proxy


def get_website_text(query):
    log.info(f"Searching: {query}")
    while True:
        url = f"https://www.google.com/search?q={query}"
        proxy = get_proxy()
        try:
            response = requests.get(url, proxies=proxy)
        except Exception as e:
            log.error(e)
            log.error("Request fail. Trying again")
            continue
        if not response.status_code == 200:
            log.error("Request fail. Trying again")
            continue
        else:
            break

    soup = BeautifulSoup(response.text, "html.parser")
    return soup.get_text()


def extract_emails(text):
    email_pattern = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b"
    emails = re.findall(email_pattern, text)
    logging.info(f"Found email id: {list(set(emails))}")
    return list(set(emails))


def save_to_excel(data, excel_file_path="output.xlsx"):
    rows = []
    for name, unique_number, emails in data:
        skip_header = False
        for email in emails:
            rows.append(
                [
                    name if not skip_header else "",
                    unique_number if not skip_header else "",
                    email,
                ]
            )
            skip_header = True
    df = pd.DataFrame(rows)
    df.to_excel(excel_file_path, index=False)


def get_final_email_ids(query):
    x = get_website_text(query)
    emails = extract_emails(x)
    if not emails:
        split_query = query.split(" ")
        if len(split_query) >= 4:
            split_query = [split_query[0], split_query[-2], split_query[-1]]
            x = get_website_text(" ".join(split_query))
            emails = extract_emails(x)
            if not emails:
                split_query.pop(0)
                x = get_website_text(" ".join(split_query))
                emails = extract_emails(x)
        elif len(split_query) == 3:
            split_query.pop(0)
            x = get_website_text(" ".join(split_query))
            emails = extract_emails(x)
    return emails


def process_row(row, result_excel_file_path):
    query = get_query(row["NAME"], row["NUMBER"])
    emails = get_final_email_ids(query)

    df = pd.DataFrame(
        {"NAME": [row["NAME"]], "NUMBER": [row["NUMBER"]], "EMAIL": [emails]}
    )
    if os.path.exists(result_excel_file_path):
        existing_df = pd.read_excel(
            result_excel_file_path, names=["NAME", "NUMBER", "EMAIL"], engine="openpyxl"
        )
        existing_df = pd.concat([existing_df, df], ignore_index=True)
        df = existing_df
    else:
        with open(result_excel_file_path, "w"):
            pass

    df = df.explode("EMAIL", ignore_index=True)
    duplicated_rows = df.duplicated(subset=["NAME", "NUMBER"])
    df.loc[duplicated_rows, ["NAME", "NUMBER"]] = ""
    log.info(f"Saved to excel: {result_excel_file_path}")

    def show_try_again_popup():
        result = messagebox.askretrycancel("Error", "Updating excel could not be possible. Please close the file if you are viewing")
        return result

    while True:
        try:
            df.to_excel(result_excel_file_path, index=False)
            break
        except:
            if not show_try_again_popup():
                continue


def get_query(name, number):
    return f"{name} {str(number).zfill(8)}"


def main():
    root = tk.Tk()
    root.geometry("400x300")
    root.title("Google Email Scraper")
    output_file_name = 'email_list.xlsx'
    data = ""
    result_excel_file_path = ""

    def choose_source_file_path():
        nonlocal data
        data = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    def choose_save_path():
        nonlocal result_excel_file_path
        result_excel_file_path = filedialog.askdirectory()
        result_excel_file_path = os.path.join(result_excel_file_path, output_file_name)

    def submit():
        nonlocal root
        if data and result_excel_file_path:
            root.destroy()
            pass
        else:
            messagebox.showerror("Error", "Please choose both source and save paths before submitting.")

    label_source = tk.Label(root, text="Please choose your source excel file:")
    label_source.pack(pady=10)

    choose_source_button = tk.Button(root, text="Choose your excel sheet", command=choose_source_file_path)
    choose_source_button.pack(pady=5)

    label_save = tk.Label(root, text="Please choose the folder to save the result Excel file:")
    label_save.pack(pady=10)

    choose_path_button = tk.Button(root, text="Choose Save Path", command=choose_save_path)
    choose_path_button.pack(pady=5)

    submit_button = tk.Button(root, text="Submit", command=submit)
    submit_button.pack(pady=20)

    root.mainloop()

    if os.path.exists(result_excel_file_path):
        base_path, extension = os.path.splitext(result_excel_file_path)
        count = 1
        result_excel_file_path = f"{base_path}({count}){extension}"

        while os.path.exists(result_excel_file_path):
            count += 1
            result_excel_file_path = f"{base_path}({count}){extension}"

    df = pd.read_excel(data, names=["NAME", "NUMBER"], engine="openpyxl")
    # Create a Toplevel window for the progress bar
    progress_window = tk.Tk()
    progress_window.title("Progress: Google Email Scraper")

    progress_frame = ttk.Frame(progress_window)
    progress_frame.pack()

    progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", length=300, mode="determinate")
    progress_bar.grid(row=0, column=0, pady=5)

    total_rows = len(df)

    for index, row in df.iterrows():
        process_row(row, result_excel_file_path=result_excel_file_path)
        progress_bar["value"] = (index + 1) * 100 / total_rows
        progress_bar.update()

    progress_window.destroy()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        log.error(e)
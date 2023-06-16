import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import logging
import gspread
import pickle
import os
import requests
from bs4 import BeautifulSoup
from typing import List, Tuple
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import openai

logging.basicConfig(filename='log.txt', level=logging.INFO)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SAMPLE_SPREADSHEET_ID = '1nSQih5Dbb4IRKUca1hnGeC_tD5m5_yhCFQFK80tXMuw'
YOUR_API_KEY = 'AIzaSyAEcWUOWgni_WtpF9-bpMXXmfKwpXdCX2Y'
YOUR_CSE_ID = '5684ee1099bbc46a2'
GPT3_API_KEY = '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'

class TDAIApplication:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("TDAI Application")

        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        self.main_frame = ttk.Frame(self.root)
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        self.table_frame = ttk.Frame(self.main_frame)
        self.table_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        self.table = ttk.Treeview(self.table_frame, columns=("1", "2", "3", "4", "5", "6"), show="headings")
        self.table.column("1", width=120)
        self.table.column("2", width=120)
        self.table.column("3", width=120)
        self.table.column("4", width=120)
        self.table.column("5", width=200)
        self.table.column("6", width=200)
        self.table.heading("1", text="Título")
        self.table.heading("2", text="Informação Extra 1")
        self.table.heading("3", text="Informação Extra 2")
        self.table.heading("4", text="Marca")
        self.table.heading("5", text="Pesquisa Google")
        self.table.heading("6", text="Dados Obtidos")
        self.table.pack(fill=tk.BOTH, expand=True)

        self.fields_frame = ttk.Frame(self.main_frame)
        self.fields_frame.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

        self.title_label = ttk.Label(self.fields_frame, text="Título 1:")
        self.title_label.pack(fill=tk.X)

        self.title_entry = ttk.Entry(self.fields_frame)
        self.title_entry.pack(fill=tk.X)

        self.title2_label = ttk.Label(self.fields_frame, text="Título 2:")
        self.title2_label.pack(fill=tk.X)

        self.title2_entry = ttk.Entry(self.fields_frame)
        self.title2_entry.pack(fill=tk.X)

        self.title3_label = ttk.Label(self.fields_frame, text="Título 3:")
        self.title3_label.pack(fill=tk.X)

        self.title3_entry = ttk.Entry(self.fields_frame)
        self.title3_entry.pack(fill=tk.X)

        self.title4_label = ttk.Label(self.fields_frame, text="Título 4:")
        self.title4_label.pack(fill=tk.X)

        self.title4_entry = ttk.Entry(self.fields_frame)
        self.title4_entry.pack(fill=tk.X)

        self.title5_label = ttk.Label(self.fields_frame, text="Título 5:")
        self.title5_label.pack(fill=tk.X)

        self.title5_entry = ttk.Entry(self.fields_frame)
        self.title5_entry.pack(fill=tk.X)

        self.description_label = ttk.Label(self.fields_frame, text="Descrição:")
        self.description_label.pack(fill=tk.X)

        self.description_text = tk.Text(self.fields_frame, height=10)
        self.description_text.pack(fill=tk.BOTH, expand=True)

        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=(0, 10), sticky="nsew")

        self.pdf_checkbox = ttk.Checkbutton(self.button_frame, text="Raspar PDF", command=self.toggle_pdf_scraping)
        self.pdf_checkbox.pack(side=tk.LEFT)
        self.pdf_path_entry = ttk.Entry(self.button_frame, width=50, state="disabled")
        self.pdf_path_entry.pack(side=tk.LEFT)
        self.pdf_browse_button = ttk.Button(self.button_frame, text="Procurar", command=self.browse_pdf_file, state="disabled")
        self.pdf_browse_button.pack(side=tk.LEFT)

        self.load_button = ttk.Button(self.button_frame, text="Carregar Dados", command=self.load_data_from_sheets)
        self.load_button.pack(side=tk.LEFT)

        self.search_button = ttk.Button(self.button_frame, text="Pesquisar no Google", command=self.search_on_google)
        self.search_button.pack(side=tk.LEFT)

        self.scrape_button = ttk.Button(self.button_frame, text="Executar Raspagem", command=self.run_scrape_script)
        self.scrape_button.pack(side=tk.LEFT)

        self.log_frame = ttk.Frame(self.main_frame)
        self.log_frame.grid(row=2, column=0, columnspan=2, padx=20, pady=(0, 10), sticky="nsew")

        self.log_label = ttk.Label(self.log_frame, text="Log:")
        self.log_label.pack(anchor=tk.W)

        self.log_text = tk.Text(self.log_frame, width=60, height=10)
        self.log_text.configure(state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.pdf_scrape = False
        self.pdf_path = ""

        self.main_frame.rowconfigure(0, weight=5)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=2)
        self.main_frame.columnconfigure(0, weight=10)
        self.main_frame.columnconfigure(1, weight=1)

        # Inicialização da API do GPT-3
        openai.api_key = GPT3_API_KEY

    def authenticate(self):
        creds = None
        token_path = "token.pickle"

        if os.path.exists(token_path):
            with open(token_path, "rb") as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentialsOAuth.json", SCOPES)
                creds = flow.run_local_server(port=0)

            with open(token_path, "wb") as token:
                pickle.dump(creds, token)

        return creds

    def loadDataFromSheets(self):
        creds = self.authenticate()
        gc = gspread.authorize(creds)
        sheet = gc.open_by_key(SAMPLE_SPREADSHEET_ID).worksheet("Data")
        values = sheet.get_all_values()
        return values

    def search_google(self, service, query, cse_id, **kwargs):
        res = service.cse().list(q=query, cx=cse_id, num=10, **kwargs).execute()
        return res['items'] if 'items' in res else None

    def searchAndPaste(self, sheet, row_num, item_name: str, item_category: str, item_third_info: str):
        query = f"{item_name} {item_category} {item_third_info}"

        service = build("customsearch", "v1", developerKey=YOUR_API_KEY)
        response = self.search_google(service, query, YOUR_CSE_ID)

        if response:
            search_results = "\n".join([item['link'] for item in response])
            sheet.update_cell(row_num, 5, search_results)

    def scrape_webpage(self, url: str) -> str:
        try:
            response = requests.get(url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, 'html.parser')
            return ' '.join([p.text for p in soup.find_all('p')])
        except Exception as e:
            logging.error(f"Error scraping webpage {url}: {str(e)}")
            return ""

    def scrape_pdf(self, pdf_path: str) -> str:
        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                text = []
                for page in reader.pages:
                    text.append(page.extract_text())
                return ' '.join(text)
        except Exception as e:
            logging.error(f"Error scraping PDF {pdf_path}: {str(e)}")
            return ""

    def save_scraped_data(self, sheet, row_num, scraped_data, link):
        # Find the next available column after column 6
        next_column = 6
        while sheet.cell(row_num, next_column).value != "":
            next_column += 1
        # Save the scraped data in the next available column
        sheet.update_cell(row_num, next_column, scraped_data)
        # Add a message to the log indicating the successful scraping for the site
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, f"Scraping successful for site in row {row_num}: {link}\n")
        self.log_text.configure(state='disabled')

    def scrape_and_return(self, sheet, row_num):
        # Get the links from cell 5 of the current row
        links = sheet.cell(row_num, 5).value.split("\n")
        scraped_content_list = []
        for link in links:
            if link:
                # Scrape each link
                scraped_content = self.scrape_webpage(link)
                scraped_content_list.append(scraped_content)
                # Add a message to the log indicating the successful scraping for the site
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, f"Scraping successful for site in row {row_num}: {link}\n")
                self.log_text.configure(state='disabled')
        # Convert the list of scraped content into a single string separated by line breaks
        scraped_data = "\n".join(scraped_content_list)
        # Save the scraped data in columns 6, 7, 8, etc.
        self.save_scraped_data(sheet, row_num, scraped_data, link)
        return scraped_data

    def add_scrape_and_paste(self, sheet, row_num):
        new_content = self.scrape_and_return(sheet, row_num)
        existing_content = sheet.cell(row_num, 6).value
        if existing_content is None:
            existing_content = ""
        sheet.update_cell(row_num, 6, existing_content + "\n" + new_content)
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, f"Added scraped content for row {row_num}\n")
        self.log_text.configure(state='disabled')

    def run_scrape_script(self):
        try:
            creds = self.authenticate()
            gc = gspread.authorize(creds)
            sheet = gc.open_by_key(SAMPLE_SPREADSHEET_ID).worksheet("Data")
            values = sheet.get_all_values()
            for i, row in enumerate(values[1:], start=2):
                self.add_scrape_and_paste(sheet, i)
            messagebox.showinfo("Success", "Scraping script ran successfully")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def load_data_from_sheets(self):
        try:
            data = self.loadDataFromSheets()
            self.display_data_in_table(data)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def search_on_google(self):
        try:
            creds = self.authenticate()
            gc = gspread.authorize(creds)
            sheet = gc.open_by_key(SAMPLE_SPREADSHEET_ID).worksheet("Data")
            values = sheet.get_all_values()
            for i, row in enumerate(values[1:], start=2):
                if row[4] == "":
                    self.searchAndPaste(sheet, i, *row[:3])
            messagebox.showinfo("Success", "Google search script ran successfully")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def display_data_in_table(self, data):
        self.table.delete(*self.table.get_children())
        for row in data[1:]:
            self.table.insert("", tk.END, values=row)

    def toggle_pdf_scraping(self):
        self.scrape_pdf = not self.scrape_pdf
        if self.scrape_pdf:
            self.pdf_path_entry.configure(state="normal")
            self.pdf_browse_button.configure(state="normal")
        else:
            self.pdf_path_entry.configure(state="disabled")
            self.pdf_browse_button.configure(state="disabled")

    def browse_pdf_file(self):
        self.pdf_path = tk.filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        self.pdf_path_entry.delete(0, tk.END)
        self.pdf_path_entry.insert(tk.END, self.pdf_path)

    def run(self):
        self.load_data_from_sheets()
        self.root.mainloop()

if __name__ == "__main__":
    app = TDAIApplication()
    app.run()

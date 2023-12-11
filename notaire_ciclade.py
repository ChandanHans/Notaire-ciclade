import os
import pickle
import json
import sys
import tkinter as tk
from threading import Thread
from time import sleep
from tkinter import filedialog
from typing import Dict, List, Tuple
from version import check_for_updates

import requests
from unidecode import unidecode
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from dotenv import load_dotenv

try:
    check_for_updates()
    print("Running the latest version.")
except Exception as e:
    print(e)
    print("\n\n!! Error !!")
    input("Press Enter to EXIT : ")
    exit()
    

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(
        os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

load_dotenv(dotenv_path=resource_path(".env"))

# File to save the settings
API_KEY = os.environ["API_KEY"]
SAVE_FILE = "settings.pkl"
SETTING_KEY = b'aY7EMKzTHYyo_gkcVoIBTUTAsWSTt2SJsbbMBwxzWsQ='
cipher_suite = Fernet(SETTING_KEY)

class ServiceAccount():
    notary_sheet_id = "1C-5OCv2Nvkr8ZSrfpnO1D5K8-kzybsu5bUa6eQL6Bj0"
    sheet_id = '1KlKBSzyFDprXy_L8Gy0UDfRfMdmpl-YZnZErg0yiATg'

    def __init__(self, email) -> None:
        self.email = email
        self.access_token = self.get_access_token()
        self.folder_id_1, self.folder_id_2 = self.get_folder_id()
        self.sheet_data = self.get_sheet_data()
        self.clients_data = self.get_clients_data()

    def get_access_token(self) -> str:
        api_url = "https://driveuploaderapi.chandanhans.repl.co/get_access_token"
        headers = {
            "api-key": API_KEY,
            'App-Identifier': 'Notaire-ciclade'
        }
        response = requests.get(api_url, headers=headers)
        data = response.json()
        return data['access_token']

    def get_folder_id(self) -> Tuple[str, str]:
        range = 'B:D'
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.notary_sheet_id}/values/{range}?majorDimension=ROWS"
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json',
        }
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()
        for row in data.get('values', []):
            if row and row[0].lower() == self.email.lower():
                return row[1], row[2]
        raise LookupError('User not found in database')

    def get_sheet_data(self) -> List[List[str]]:
        # Fetch the fifth row first to find the columns of interest
        fifth_row_range = 'Suivi factures!5:5'
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.sheet_id}/values:batchGet"
        params = {
            'ranges': [fifth_row_range],
            'majorDimension': 'ROWS',
        }
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json',
        }
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        fifth_row = response.json()['valueRanges'][0].get('values', [[]])[0]

        # Identify columns for 'Date de naissance' and 'Date de mort'
        dob_col = None
        dod_col = None
        name_col = None
        for index, value in enumerate(fifth_row):
            if value == "Nom/Prénom":
                name_col = index
            elif value == 'Date de naissance':
                dob_col = index
            elif value == 'Date de mort':
                dod_col = index
            
        def number_to_column(n):
            string = ""
            n += 1
            while n > 0:
                n, remainder = divmod(n - 1, 26)
                string = chr(65 + remainder) + string
            return string
        
        dob_col_letter = number_to_column(dob_col)
        dod_col_letter = number_to_column(dod_col)
        name_col_letter = number_to_column(name_col)
        
        # Now fetch these columns
        ranges = [f'Suivi factures!{name_col_letter}:{name_col_letter}',
                  f'Suivi factures!{dob_col_letter}:{dob_col_letter}',
                  f'Suivi factures!{dod_col_letter}:{dod_col_letter}']
        params = {
            'ranges': ranges,
            'majorDimension': 'COLUMNS',
        }
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        columns_data = response.json()['valueRanges']
        sheet_data = []
        max_length = max(len(column_data.get('values', [[]])[0]) for column_data in columns_data if column_data.get('values'))
        for i in range(max_length):
            row_data = []
            skip_row = False
            for column_data in columns_data:
                column_values = column_data.get('values', [[]])[0]
                value = column_values[i] if i < len(column_values) else ''
                if not value:
                    skip_row = True
                    break
                row_data.append(value)
            if not skip_row:
                sheet_data.append(row_data)
        return sheet_data

    def get_target_folders(self):
        url = "https://www.googleapis.com/drive/v3/files"
        headers = {
            'Authorization': f'Bearer {self.access_token}'            
        }
        params = {
            'q': f"'{self.folder_id_1}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            'pageSize': 100,
            'fields': 'nextPageToken, files(id, name)'
        }
        folders = []
        while True:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            file_list = response.json().get('files', [])
            folders.extend([{'id': folder['id'], 'name': folder['name']}
                           for folder in file_list])
            page_token = response.json().get('nextPageToken', None)
            if page_token is None:
                break
            else:
                params['pageToken'] = page_token
        return folders

    def get_death_proof(self, client_data) -> Dict[str, str]:
        url = "https://www.googleapis.com/drive/v3/files"
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }
        params = {
            'q': f"name contains 'acte de d' and '{client_data['id']}' in parents and trashed=false",
            'pageSize': 10,
            'fields': 'files(id, name)'
        }
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        file_list = response.json().get('files')
        if file_list:
            return {'id': file_list[0]['id'], 'name': file_list[0]['name']}

    def get_mandat(self, client_data) -> Dict[str, str]:
        url = "https://www.googleapis.com/drive/v3/files"
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }
        params = {
            'q': f"name contains 'Mandat' and '{client_data['id']}' in parents and trashed=false",
            'pageSize': 10,
            'fields': 'files(id, name)'
        }
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        file_list = response.json().get('files')
        if file_list:
            return {'id': file_list[0]['id'], 'name': file_list[0]['name']}

    def download_file(self, file: Dict[str, str]) -> str:
        file_id = file['id']
        file_name = file['name']
        headers = {
            "Authorization": f"Bearer {self.access_token}"
        }
        download_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
        response = requests.get(download_url, headers=headers, stream=True)
        response.raise_for_status()

        with open(file_name, 'wb') as f:
            for chunk in response.iter_content(32768):
                f.write(chunk)
        return file_name

    def get_client_data(self, folder: dict[str, str]) -> Dict:
        result = folder.copy()
        for row in self.sheet_data:
            # Assuming the folder name is in the first column
            if row and unidecode(row[0].lower().strip()) == unidecode(folder['name'].lower().split("(")[0].strip()):
                result['fname'], result['lname'] = self.split_name(row[0])
                result['dob'] = row[-2]
                result['dod'] = row[-1]
                result['death_proof'] = self.get_death_proof(folder)
                result['mandat'] = self.get_mandat(folder)
                return result

    def get_clients_data(self) -> List:
        clients_data = []
        target_folders = self.get_target_folders()
        for folder in target_folders:
            client_data = self.get_client_data(folder)
            if client_data and client_data['death_proof'] and client_data['mandat']:
                clients_data.append(client_data)
        return clients_data

    def move_folder(self, client_data):
        folder_id_to_move = client_data["id"]
        patch_url = f'https://www.googleapis.com/drive/v3/files/{folder_id_to_move}'
        headers = {
            'Authorization': f'Bearer {self.access_token}'
        }
        params = {
            'addParents': self.folder_id_2,
            'removeParents': self.folder_id_1
        }
        patch_response = requests.patch(patch_url, headers=headers, params=params)
        patch_response.raise_for_status()
        print(f"Folder moved successfully")
    
    @staticmethod
    def split_name(full_name: str) -> Tuple[str, str]:
        name_parts = full_name.split()
        last_name_parts = [part for part in name_parts if part.isupper()]
        first_name_parts = [part for part in name_parts if not part.isupper()]
        last_name = " ".join(last_name_parts)
        first_name = " ".join(first_name_parts)
        return last_name, first_name


def load_settings():
    if os.path.exists(SAVE_FILE):
        with open(SAVE_FILE, 'rb') as f:
            encrypted_data = f.read()
            decrypted_data = cipher_suite.decrypt(encrypted_data)
            return pickle.loads(decrypted_data)
    return {}

def save_settings():
    settings = {
        "email": email_var.get(),
        "password": password_var.get(),
        "owner": owner_var.get(),
        "iban": iban_var.get(),
        "bic": bic_var.get(),
        "pdfPath": pdf_file_path
    }
    encrypted_data = cipher_suite.encrypt(pickle.dumps(settings))
    with open(SAVE_FILE, 'wb') as f:
        f.write(encrypted_data)


def choose_file():
    global pdf_file_path
    filepath = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_file_path = filepath
    pdf_var.set(os.path.basename(filepath))



def start_browser() -> webdriver.Chrome:
    extension_path = resource_path('CapSolver')
    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument("--disable-popup-blocking")
    options.add_experimental_option("prefs", {
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
        "autofill.credit_card_enabled": False,  # Disable Chrome's credit card autofill
        "autofill.enabled": False,  # Disable all autofill
    })
    options.add_argument(f'--load-extension={extension_path}')
    options.add_argument("--app=https://ciclade.caissedesdepots.fr/monespace")
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(15)
    try:
        driver.find_element(
            By.XPATH, '//*[@id="didomi-notice-agree-button"]').click()
        sleep(3)
    except Exception as e:
        print(f"Error while agreeing to terms: {e}")
    login(driver)
    return driver

def solve_captcha(driver: webdriver.Chrome):
    solver_element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="capsolver-solver-info"]')))
    while solver_element.text != "Captcha solved!":
        solver_element = WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH,'//div[@class="capsolver-solver-info"]')))
        sleep(2)
        
        
def login(driver: webdriver.Chrome):
    driver.find_element(
        By.XPATH, '//*[@id="login"]').send_keys(email_var.get())
    driver.find_element(
        By.XPATH, '//*[@id="f-login-passw"]').send_keys(password_var.get())
    driver.find_element(By.XPATH, '//button[@type="submit"]').click()
    driver.find_element(By.XPATH, '//*[@class="ttl-is-h1 ng-binding"]')

def click_element(driver : webdriver.Chrome, xpath):
    """Utility function to click on an element identified by xpath"""
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
    

def send_keys_to_element(driver : webdriver.Chrome, xpath, text):
    """Utility function to send text to an element identified by xpath"""
    WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath))).send_keys(text)


def upload_to_element(driver : webdriver.Chrome, xpath, path):
    """Utility function to send text to an element identified by xpath"""
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))).send_keys(path)

def new_search(driver : webdriver.Chrome, client_data : Dict, temp_file1, temp_file2) -> bool :
    try:
        driver.get("https://ciclade.caissedesdepots.fr/monespace/#/service/recherche")
        driver.refresh()

        # Fill out the form with the client data
        click_element(driver, '//input[@id="f-s-p-death-yes"]')
        click_element(driver, '//*[@id="f-s-p-civilite-monsieur"]')
        send_keys_to_element(driver, '//input[@id="f-s-p-death-day"]', client_data["dod"])
        send_keys_to_element(driver, '//input[@id="f-s-p-birth-day"]', client_data["dob"])
        send_keys_to_element(driver, '//input[@id="f-s-p-surname1"]', client_data["fname"])
        send_keys_to_element(driver, '//input[@id="f-s-p-name1"]', client_data["lname"])
        click_element(driver, '//*[@id="f-s-p-nationality"]/option[@value="FRA"]')
        click_element(driver, '//*[@id="f-s-p-connu-no"]')
        solve_captcha(driver)
        sleep(1)
        click_element(driver, '//button[@type="submit"]')
        click_element(driver, '//button[@ng-click="vm.rechercher()"]')
        click_element(driver, '//*[@ng-if="vm.adresseFiscaleRenseigne"]')
        
        # Step 1
        for _ in range(5):
            try:
                WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="f-s-p-titulaire"]')))
                sleep(3)
                click_element(driver, '//*[@id="positionDemandeur"]/option[@label="Notaire"]')
                click_element(driver, '//*[@id="f-s-p-paysBanque"]/option[@value="FR"]')
                send_keys_to_element(driver, '//*[@id="f-s-p-titulaire"]', owner_var.get())
                send_keys_to_element(driver, '//*[@id="f-s-p-iban"]', iban_var.get())
                send_keys_to_element(driver, '//*[@id="f-s-p-bic"]', bic_var.get())
                upload_to_element(driver, '//*[@id="document"]', pdf_file_path)
                click_element(driver, '//*[@id="page_top"]/div[2]/p[2]/button')
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="docAdditionnelNon"]')))
                break
            except:
                if _ == 4:
                    raise
                print("Error in Step 1")
                sleep(5)
                pass
        
        # Step 2
        for _ in range(5):
            try:
                element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="docAdditionnelNon"]')))
                sleep(3)
                element.click()
                upload_to_element(driver, '//*[@id="document-0"]', temp_file1)
                upload_to_element(driver, '//*[@id="document-1"]', temp_file2)
                click_element(driver, '//*[@id="page_top"]/div[2]/p[5]/button')
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSoumission"]')))
                break
            except:
                if _ == 4:
                    raise
                print("Error in Step 2")
                sleep(5)
                driver.refresh()
                pass
        
        # Final submission
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="btnSoumission"]')))
        # sleep(3)
        # submit_button.click()
        sleep(5)
        return True
    except Exception as e:
        print(f"An error occurred: {e}")
        return False
    
    

def start_process():
    start_btn.config(state=tk.DISABLED)
    email = email_var.get()
    try:
        status_label.config(text="Getting information from Drive.", fg="#4CAF50")
        sa = ServiceAccount(email)
        status_label.config(text="Running...", fg="#4CAF50")
        driver = start_browser()
        clients_data = sa.clients_data
        for index,client_data in enumerate(clients_data):
            status_label.config(text=f'{index}/{len(clients_data)}\n\n{client_data["name"]}', fg="#4CAF50")
            file1_name = sa.download_file(client_data["death_proof"])
            file2_name = sa.download_file(client_data["mandat"])
            file1_path = os.path.join(os.getcwd(), file1_name)
            file2_path = os.path.join(os.getcwd(), file2_name)
            successful = new_search(driver, client_data, file1_path, file2_path)
            os.remove(file1_name)
            os.remove(file2_name)
            if successful:
                sa.move_folder(client_data)
        status_label.config(text="Completed.", fg="#4CAF50")
    except requests.ConnectionError:
        status_label.config(text=f"No internet connection!", fg="#FF0000")
    except LookupError as e:
        status_label.config(text=f"Contact KLERO to add your information!", fg="#FF0000")
    except Exception as e:
        status_label.config(text=f"Error: {e}", fg="#FF0000")
    finally:
        start_btn.config(state=tk.NORMAL)


def start():
    if (owner_var and iban_var and bic_var and pdf_file_path):
        if (os.path.exists(pdf_file_path)):
            save_settings()
            thread = Thread(target=start_process)
            thread.start()
        else:
            status_label.config(text=f"Select RIB Pdf again!", fg="#FF0000")
    else:
        status_label.config(text=f"Fill all the information!", fg="#FF0000")



root = tk.Tk()
root.title("Notaire Ciclade Test")
root.iconbitmap(resource_path('icon.ico'))
root.geometry("600x500")
root.configure(bg='#2c2f33')
root.resizable(False, False)  # Fix window size

# Styling
BUTTON_STYLE = {
    'bg': '#3B3E44',
    'fg': '#ffffff',
    'activebackground': '#464B51',
    'border': '0',
    'font': ('Arial', 12, 'bold')
}

LABEL_STYLE = {
    'bg': '#2c2f33',
    'fg': '#ffffff',
    'font': ('Arial', 14, 'bold')
}

ENTRY_STYLE = {
    'bg': '#383c41',
    'fg': '#ffffff',
    'borderwidth': '1px',
    'relief': 'solid',
    'width': 30,
    'font': ('Arial', 12)
}

STATUS_LABEL_STYLE = {
    'bg': '#2c2f33',
    'font': ('Arial', 12, 'bold')
}

# Load previous settings
settings = load_settings()
pdf_file_path = settings.get("pdfPath", "")

# Variables
email_var = tk.StringVar(value=settings.get("email", ""))
password_var = tk.StringVar(value=settings.get("password", ""))
owner_var = tk.StringVar(value=settings.get("owner", ""))
iban_var = tk.StringVar(value=settings.get("iban", ""))
bic_var = tk.StringVar(value=settings.get("bic", ""))
pdf_var = tk.StringVar(value=os.path.basename(pdf_file_path))

# Layout
rows = [
    ("Email:", email_var),
    ("Password:", password_var),
    ("Account owner:", owner_var),
    ("IBAN:", iban_var),
    ("BIC:", bic_var)
]

for idx, (text, var) in enumerate(rows, start=1):
    tk.Label(root, text=text, **LABEL_STYLE).grid(row=idx,
                                                  column=0, sticky='w', padx=30, pady=5)
    if text == "Password:":
        # If this is the password field, use the show parameter to mask the input
        entry = tk.Entry(root, textvariable=var, show='*', **ENTRY_STYLE)
    else:
        entry = tk.Entry(root, textvariable=var, **ENTRY_STYLE)
    entry.grid(row=idx, column=1, padx=30, pady=5, sticky='e')

label = tk.Label(root, text="RIB PDF:", **LABEL_STYLE)
label.grid(row=len(rows) + 1, column=0, padx=30, pady=5, sticky='w')
pdf_frame = tk.Frame(root, bg='#2c2f33')
pdf_frame.grid(row=len(rows) + 1, column=1, padx=30, pady=5, sticky='e')
tk.Label(pdf_frame, textvariable=pdf_var, bg="#2c2f33", fg="#ffffff",
         width=33, anchor='e', font=('Arial', 12)).pack(side=tk.LEFT)
tk.Button(pdf_frame, text="...", command=lambda: choose_file(),
          **BUTTON_STYLE).pack(side=tk.RIGHT)


start_btn = tk.Button(root, text="Start Browser",
                      command=start, **BUTTON_STYLE)
start_btn.grid(row=len(rows) + 5, column=0, columnspan=2, pady=10)
status_label = tk.Label(
    root, text="", **STATUS_LABEL_STYLE, wraplength=500, anchor="w")
status_label.grid(row=len(rows) + 6, column=0, columnspan=2, pady=10)

root.mainloop()

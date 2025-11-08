import requests
import json
import uuid
import os
from datetime import datetime
import rsa
import openpyxl
import csv
from tqdm import tqdm

# -------- Paths / "original intended folder" behavior --------
# By default, save right next to the script (same as your original behavior).
# If you prefer the current working directory, change BASE_DIR to os.getcwd().
try:
    BASE_DIR = os.path.abspath(os.path.dirname(__file__))
except NameError:
    BASE_DIR = os.getcwd()  # fallback if running in an environment without __file__

def _join_here(*parts) -> str:
    return os.path.join(BASE_DIR, *parts)

def _ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

# -------- Date parsing (robust) --------
def _safe_date(date_str: str) -> str:
    """
    Return YYYY-MM-DD from various Tricount date string shapes.
    Tries microseconds, then seconds, then fromisoformat. Falls back to original string.
    """
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(date_str).strftime("%Y-%m-%d")
    except Exception:
        return date_str

class TricountAPI:
    def __init__(self):
        self.base_url = "https://api.tricount.bunq.com"
        self.app_installation_id = str(uuid.uuid4())
        self.public_key, self.private_key = rsa.newkeys(2048)
        self.rsa_public_key_pem = self.public_key.save_pkcs1(format="PEM").decode()
        self.headers = {
            "User-Agent": "com.bunq.tricount.android:RELEASE:7.0.7:3174:ANDROID:13:C",
            "app-id": self.app_installation_id,
            "X-Bunq-Client-Request-Id": "049bfcdf-6ae4-4cee-af7b-45da31ea85d0"
        }
        self.auth_token = None
        self.user_id = None

    def authenticate(self):
        auth_url = f"{self.base_url}/v1/session-registry-installation"
        auth_payload = {
            "app_installation_uuid": self.app_installation_id,
            "client_public_key": self.rsa_public_key_pem,
            "device_description": "Android"
        }
        response = requests.post(auth_url, json=auth_payload, headers=self.headers)
        response.raise_for_status()
        auth_data = response.json()

        response_items = auth_data["Response"]
        self.auth_token = next(item["Token"]["token"] for item in response_items if "Token" in item)
        self.user_id = next(item["UserPerson"]["id"] for item in response_items if "UserPerson" in item)
        self.headers["X-Bunq-Client-Authentication"] = self.auth_token

    def fetch_tricount_data(self, tricount_key):
        tricount_url = f"{self.base_url}/v1/user/{self.user_id}/registry?public_identifier_token={tricount_key}"
        response = requests.get(tricount_url, headers=self.headers)
        response.raise_for_status()
        return response.json()

class TricountHandler:
    # ---------- Parsing / Order ----------
    @staticmethod
    def get_tricount_title(data):
        return data["Response"][0]["Registry"]["title"]

    @staticmethod
    def parse_tricount_data(data):
        """
        Preserve original order:
        - memberships in the order provided by the API
        - transactions in the order provided by the API
        """
        registry = data["Response"][0]["Registry"]

        memberships = [
            {"Name": m["RegistryMembershipNonUser"]["alias"]["display_name"]}
            for m in registry["memberships"]
        ]  # DO NOT sort; keep original API order

        transactions = []
        for entry in registry["all_registry_entry"]:  # original order
            transaction = entry["RegistryEntry"]
            type_transaction = transaction["type_transaction"]
            who_paid = transaction["membership_owned"]["RegistryMembershipNonUser"]["alias"]["display_name"]
            total = float(transaction["amount"]["value"]) * -1
            currency = transaction["amount"]["currency"]
            description = transaction.get("description", "")
            when = transaction["date"]
            shares = {
                alloc["membership"]["RegistryMembershipNonUser"]["alias"]["display_name"]: abs(float(alloc["amount"]["value"]))
                for alloc in transaction["allocations"]
            }
            category = transaction["category"]
            attachments = transaction.get("attachment", [])

            transactions.append({
                "Type": type_transaction,
                "Who Paid": who_paid,
                "Total": total,
                "Currency": currency,
                "Description": description,
                "When": when,
                "Shares": shares,
                "Category": category,
                "Attachments": attachments
            })

        return memberships, transactions

    # ---------- Attachments ----------
    @staticmethod
    def download_attachments(transactions, download_folder_name):
        download_folder = _join_here(download_folder_name)
        _ensure_dir(download_folder)

        file_counter = 1
        total_files = sum(len(transaction["Attachments"]) for transaction in transactions)
        print(f"Total Attachments: {total_files}")

        if total_files == 0:
            return

        with tqdm(total=total_files, desc="Downloading attachments") as progress_bar:
            for transaction in transactions:
                attachment_files = []
                for attach in transaction["Attachments"]:
                    if "urls" in attach and attach["urls"]:
                        url = attach["urls"][0]["url"]
                        extension = os.path.splitext(url.split("?")[0])[1] or ".file"
                        file_name = f"receipt_{file_counter}{extension}"
                        file_path = os.path.join(download_folder, file_name)
                        TricountHandler._download_file(url, file_path)
                        attachment_files.append(file_name)
                        file_counter += 1
                        progress_bar.update(1)
                transaction["File Names"] = ", ".join(attachment_files)

    @staticmethod
    def _download_file(url, file_path):
        response = requests.get(url)
        response.raise_for_status()
        with open(file_path, "wb") as file:
            file.write(response.content)

    # ---------- Row preparation (order preserved) ----------
    @staticmethod
    def prepare_transaction_data(transaction):
        """
        Columns (fixed order):
        Who Paid | Total | Currency | Description | When | Involved | File Names | Attachment URLs | Category
        """
        involved = ", ".join([name for name, amount in transaction["Shares"].items() if amount > 0])
        row_data = [
            transaction["Who Paid"],
            transaction["Total"],
            transaction["Currency"],
            transaction["Description"],
            _safe_date(transaction["When"]),
            involved,
            transaction.get("File Names", ""),
            ", ".join([attach["urls"][0]["url"] for attach in transaction["Attachments"] if "urls" in attach and attach["urls"]]),
            transaction["Category"]
        ]
        return row_data

    @staticmethod
    def prepare_sesterce_transaction_data(transaction, members_in_original_order):
        """
        A row contains:
        Date, Title,
        Paid by Member A..N (original membership order),
        Paid for Member A..N (original membership order),
        Currency, Category
        """
        members = members_in_original_order  # keep original order

        paid_by = [0.0] * len(members)
        payer = transaction["Who Paid"]
        if payer in members:
            paid_by[members.index(payer)] = transaction["Total"]

        paid_for = [0.0] * len(members)
        for paid_for_member, amount in transaction["Shares"].items():
            if paid_for_member in members:
                paid_for[members.index(paid_for_member)] = amount

        type_transaction = transaction["Type"]
        category = ""
        if type_transaction == "BALANCE":
            category = "Money Transfer"
        elif type_transaction == "INCOME":
            paid_for = [-amount for amount in paid_for]
            category = transaction["Category"] if transaction["Category"] != "UNCATEGORIZED" else ""
        elif type_transaction == "NORMAL":
            category = transaction["Category"] if transaction["Category"] != "UNCATEGORIZED" else ""

        row_data = [
            _safe_date(transaction["When"]),
            transaction["Description"],
            *paid_by,
            *paid_for,
            transaction["Currency"],
            category
        ]
        return row_data

    # ---------- Writers (save next to script; keep original filenames) ----------
    @staticmethod
    def write_to_excel(transactions, file_name):
        """
        Writes to {BASE_DIR}/{file_name}.xlsx
        """
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Tricount Transactions"

        headers = ["Who Paid", "Total", "Currency", "Description", "When", "Involved", "File Names", "Attachment URLs", "Category"]
        sheet.append(headers)

        for transaction in transactions:
            row_data = TricountHandler.prepare_transaction_data(transaction)
            sheet.append(row_data)

        xlsx_path = _join_here(f"{file_name}.xlsx")
        workbook.save(xlsx_path)
        print(f"Transactions have been saved to {xlsx_path}.")

    @staticmethod
    def write_to_csv(transactions, file_name):
        """
        Semicolon-delimited CSV (same as original), UTF-8 with BOM so Excel shows Unicode correctly.
        Saves to {BASE_DIR}/{file_name}.csv
        """
        csv_path = _join_here(f"{file_name}.csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as csvfile:
            headers = ["Who Paid", "Total", "Currency", "Description", "When", "Involved", "File Names", "Attachment URLs", "Category"]
            writer = csv.writer(csvfile, delimiter=";")
            writer.writerow(headers)
            for transaction in transactions:
                row_data = TricountHandler.prepare_transaction_data(transaction)
                writer.writerow(row_data)
        print(f"Transactions have been saved to {csv_path}.")

    @staticmethod
    def write_to_sesterce_csv(memberships, transactions, file_name):
        """
        Comma-delimited, UTF-8 with BOM, **original member order** (no sorting).
        Saves to {BASE_DIR}/{file_name}.csv
        """
        members_in_original_order = [member["Name"] for member in memberships]  # DO NOT sort

        csv_path = _join_here(f"{file_name}.csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as csvfile:
            headers = (
                ["Date", "Title"]
                + [f"Paid by {m}" for m in members_in_original_order]
                + [f"Paid for {m}" for m in members_in_original_order]
                + ["Currency", "Category"]
            )
            writer = csv.writer(csvfile, delimiter=",")
            writer.writerow(headers)
            for transaction in transactions:
                row_data = TricountHandler.prepare_sesterce_transaction_data(transaction, members_in_original_order)
                writer.writerow(row_data)
        print(f"Transactions have been saved to {csv_path}.")

if __name__ == "__main__":
    # example key (replace with yours)
    tricount_key = "tISWyMCgrIMgFuxudZ"

    api = TricountAPI()
    api.authenticate()
    data = api.fetch_tricount_data(tricount_key)

    # save data to local file next to the script (original intended folder)
    response_json_path = _join_here("response_data.json")
    with open(response_json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Wrote raw JSON to {response_json_path}")

    handler = TricountHandler()
    tricount_title = handler.get_tricount_title(data)

    memberships, transactions = handler.parse_tricount_data(data)

    # CSV in original order, saved next to the script
    handler.write_to_csv(transactions, file_name=f"Transactions {tricount_title}")

    # Optional extras (same folder & original names):
    # handler.write_to_excel(transactions, file_name=f"Transactions {tricount_title}")
    # handler.write_to_sesterce_csv(memberships, transactions, f"Transaction {tricount_title} (Sesterce)")
    # handler.download_attachments(transactions, download_folder_name=f"Attachments {tricount_title}")

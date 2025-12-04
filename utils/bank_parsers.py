import pdfplumber
import camelot
import re
import pandas as pd
from datetime import datetime
import streamlit as st

# --- Helper Functions ---

def parse_td_visa_amount(amount_str):
    """Cleans and converts VISA amount string to a float, handling 'CR'."""
    if not isinstance(amount_str, str): return 0.0
    is_credit = 'CR' in amount_str
    cleaned_str = re.sub(r'[^\d.]', '', amount_str)
    if not cleaned_str: return 0.0
    amount = float(cleaned_str)
    return amount if is_credit else -amount

def get_statement_year(pdf_file_object):
    """Extracts the closing year from the first page of a statement."""
    try:
        with pdfplumber.open(pdf_file_object) as pdf:
            pdf_file_object.seek(0)
            page1_text = pdf.pages[0].extract_text() or ""
            
            match = re.search(r'\w+\s+\d{1,2},\s+\d{4}\s+-\s+\w+\s+\d{1,2},\s+(\d{4})', page1_text)
            if match: return int(match.group(1))
            
            match = re.search(r'Statement Period:\s*\w+\s\d{1,2}\s(\d{4})', page1_text)
            if match: return int(match.group(1))
            
            match = re.search(r'\w+\s+\d{1,2},\s+(\d{4})\s+to\s+\w+\s+\d{1,2},\s+\d{4}', page1_text)
            if match: return int(match.group(1))
            
        return datetime.now().year
    except Exception:
        return datetime.now().year

# --- Bank Parsers ---

def parse_bank_of_america(file_object):
    st.toast(f"Processing Bank of America: {file_object.name}...")
    transactions = []
    try:
        with pdfplumber.open(file_object) as pdf:
            file_object.seek(0)
            full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

        # Withdrawals
        try:
            withdrawals_block = re.search(r'Withdrawals and other debits\n(.*?)\nTotal withdrawals and other debits', full_text, re.DOTALL).group(1)
            current_transaction = None
            for line in withdrawals_block.strip().split('\n'):
                match = re.match(r'^(\d{2}/\d{2}/\d{2})\s+(.*?)\s+(-?[\d,]+\.\d{2})$', line)
                if match:
                    if current_transaction:
                        full_date = datetime.strptime(current_transaction['date'], "%m/%d/%y").strftime("%Y-%m-%d")
                        amount = float(re.sub(r'[^\d.-]', '', current_transaction['amount']))
                        desc = ' '.join(current_transaction['desc_parts'])
                        transactions.append(("Bank of America", full_date, "", desc, amount))
                    date_str, desc_part, amount_str = match.groups()
                    current_transaction = {'date': date_str, 'desc_parts': [desc_part.strip()], 'amount': amount_str}
                elif current_transaction:
                    current_transaction['desc_parts'].append(line.strip())
            if current_transaction:
                full_date = datetime.strptime(current_transaction['date'], "%m/%d/%y").strftime("%Y-%m-%d")
                amount = float(re.sub(r'[^\d.-]', '', current_transaction['amount']))
                desc = ' '.join(current_transaction['desc_parts'])
                transactions.append(("Bank of America", full_date, "", desc, amount))
        except AttributeError:
            pass 

        # Fees
        try:
            fees_block = re.search(r'Service fees - continued\n(.*?)\nTotal service fees', full_text, re.DOTALL).group(1)
            matches = re.findall(r'(\d{2}/\d{2}/\d{2})\s+(.*?)\s+(-?[\d,]+\.\d{2})', fees_block)
            for date_str, desc, amount_str in matches:
                full_date = datetime.strptime(date_str, "%m/%d/%y").strftime("%Y-%m-%d")
                amount = float(re.sub(r'[^\d.-]', '', amount_str))
                if amount != 0:
                    transactions.append(("Bank of America", full_date, "", ' '.join(desc.split()), amount))
        except AttributeError:
            pass

    except Exception as e:
        st.error(f"Error parsing BoA {file_object.name}: {e}")
    return transactions

def parse_td_visa_card(file_object):
    st.toast(f"Processing TD VISA: {file_object.name}...")
    transactions = []
    year = get_statement_year(file_object)
    try:
        tables = camelot.read_pdf(file_object, pages='3', flavor='stream')
        if not tables: return []
        df = tables[0].df.replace(r'^\s*$', float('nan'), regex=True)
        
        header_row = -1
        for i, row in df.iterrows():
             row_text = ' '.join(str(s) for s in row if pd.notna(s))
             if 'Activity Date' in row_text and 'Reference Number' in row_text:
                 header_row = i
                 break
        if header_row == -1: return []

        for i in range(header_row + 1, len(df)):
            row = df.iloc[i]
            if pd.notna(row[1]) and pd.notna(row.iloc[-1]):
                full_date_str = f"{row[1].strip()} {year}"
                try:
                    date_obj = datetime.strptime(full_date_str, "%b %d %Y").strftime("%Y-%m-%d")
                    amount = parse_td_visa_amount(row.iloc[-1])
                    ref = row[2]
                    desc = ' '.join(str(s) for s in row[3:-1] if pd.notna(s))
                    transactions.append(('TD BUSINESS SOLUTIONS VISA', date_obj, ref, desc, amount))
                except:
                    continue
    except Exception as e:
        st.error(f"Error parsing TD VISA {file_object.name}: {e}")
    return transactions

def parse_td_generic(file_object, bank_name, credit_headers, debit_headers):
    st.toast(f"Processing {bank_name}: {file_object.name}...")
    transactions = []
    year = get_statement_year(file_object)
    all_headers = credit_headers + debit_headers
    try:
        with pdfplumber.open(file_object) as pdf:
            file_object.seek(0)
            page = pdf.pages[0]
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""

            in_section, current_type = False, None

            for line in text.split('\n'):
                line = line.strip()
                if not line: continue

                matched = False
                for h in all_headers:
                    if line.startswith(h):
                        in_section, current_type, matched = True, 'credit' if h in credit_headers else 'debit', True
                        break
                if matched: continue

                if line.startswith("Subtotal:"):
                    in_section = False
                    continue
                if in_section and "POSTING DATE" in line: continue

                if in_section:
                    match = re.match(r'^(\d{2}/\d{2})\s+(.*?)\s+([\d,]+\.\d{2})$', line)
                    if match:
                        d_str, desc, amt_str = match.groups()
                        full_date = datetime.strptime(f"{d_str}/{year}", "%m/%d/%Y").strftime("%Y-%m-%d")
                        amt = float(re.sub(r'[^\d.]', '', amt_str))
                        if current_type == 'debit': amt = -amt
                        transactions.append((bank_name, full_date, '', desc.strip(), amt))
    except Exception as e:
        st.error(f"Error parsing {bank_name} ({file_object.name}): {e}")
    return transactions
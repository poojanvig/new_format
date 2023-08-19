# dont forget to uplad the "axis_cateegory.xlsx" sheet
import re
import os
import io
import pandas as pd
import numpy as np
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import pdfplumber
from datetime import datetime
from dateutil.relativedelta import relativedelta
import regex as re
import datefinder
from calendar import monthrange
import calendar
from openpyxl.styles import Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

pd.options.display.float_format = "{:,.2f}".format


class SingleBankStatementConverter:
    def __init__(self, bank_names, pdf_paths, pdf_passwords, start_date, end_date, account_number, file_name):
        self.writer = None
        self.bank_names = bank_names
        self.pdf_paths = pdf_paths
        self.pdf_passwords = pdf_passwords
        self.start_date = start_date
        self.end_date = end_date
        self.account_number = account_number
        self.file_name = None

    @classmethod
    def extract_text_from_pdf(cls, filename):
        # Open the PDF file in read-binary mode
        with open(filename, "rb") as file:
            # Create a PDF file reader object
            pdf_reader = PyPDF2.PdfFileReader(file)

            # Get the text from the first page
            page = pdf_reader.getPage(0)
            text = page.extractText()
            lines = text.split("\n")

            # Extract Account Name
            match = re.search(r"Account Name\s+:\s+(.*)", text)
            if match:
                account_name = match.group(1).strip()
            else:
                account_name = "Account Name not found"

            print(account_name)
            # Extract Account Number
            match = re.search(r"Account Number\s+:\s+(\d+)", text)
            if match:
                account_number = match.group(1)
                cls.masked_account_number = (
                        "*" * (len(account_number) - 4) + account_number[-4:]
                )
            else:
                account_number = None
                cls.masked_account_number = "Account Number not found"
            print("\n\nThis is the account number ", account_number, "\n\n")

            # Account desc
            match1 = re.search(r"Account Description\s*:\s*(.*)", text)

            if match1:
                cls.account_description = match1.group(1)
                print(cls.account_description)
            else:
                print("Account Description not found")

            return account_number

    def unlock_the_pdfs_path(self, pdf_path, pdf_password):
        # Create the "saved_pdf" folder if it doesn't exist
        os.makedirs("saved_pdf", exist_ok=True)
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            if pdf_reader.is_encrypted:
                pdf_reader.decrypt(pdf_password)
                try:
                    _ = pdf_reader.numPages  # Check if decryption was successful
                    pdf_writer = PyPDF2.PdfWriter()
                    for page in pdf_reader.pages:
                        pdf_writer.add_page(page)
                    unlocked_pdf_path = f"saved_pdf/unlocked.pdf"
                    with open(unlocked_pdf_path, 'wb') as unlocked_pdf_file:
                        pdf_writer.write(unlocked_pdf_file)
                    print("PDF unlocked and saved successfully.")
                except PyPDF2.utils.PdfReadError:
                    print("Incorrect password. Unable to unlock the PDF.")
            else:
                # Copy the PDF file to the "saved_pdf" folder without modification
                unlocked_pdf_path = f"saved_pdf/unlocked.pdf"
                with open(pdf_path, 'rb') as unlocked_pdf_file:
                    with open(unlocked_pdf_path, 'wb') as output_file:
                        output_file.write(unlocked_pdf_file.read())
                print("PDF saved in the 'saved_pdf' folder.")
        return unlocked_pdf_path

    def insert_separator(self, page, y_position, page_width, page_height):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)  # Set fill color to black
        can.setStrokeColorRGB(0, 0, 0)  # Set stroke color to black
        can.setLineWidth(2)  # Set line width
        line_height = page_height / (len(page.extract_text().split('\n')) + 1)
        line_extension = page_width * 2
        can.line(0, y_position, page_width + line_extension, y_position)
        can.save()
        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page_width + line_extension, page_height)
        page.merge_page(sep_page)

    def separate_lines_in_pdf(self, input_pdf_path):
        output_pdf_path = 'saved_pdf/output_horizontal.pdf'
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()
        for page_num in range(len(input_pdf.pages)):
            page = input_pdf.pages[page_num]
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]
            content = page.extract_text()
            lines = content.split('\n')
            for line_num, line in enumerate(lines):
                if line.strip():
                    self.insert_separator(page, line_num * (page_height / (len(lines) + 1)), page_width, page_height)
            output_pdf.add_page(page)
        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)
        return output_pdf_path

    def insert_vertical_lines(self, page, x_positions, page_height):
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        can.setFillColorRGB(0, 0, 0)  # Set fill color to black
        can.setStrokeColorRGB(0, 0, 0)  # Set stroke color to black
        can.setLineWidth(2)  # Set line width
        for x_position in x_positions:
            can.line(x_position, 0, x_position, page_height)
        can.save()
        packet.seek(0)
        sep_pdf = PdfReader(packet)
        sep_page = sep_pdf.pages[0]
        sep_page.mediabox.upper_right = (page.mediabox.upper_right[0], page_height)
        page.merge_page(sep_page)

    def separate_lines_in_vertical_pdf(self, input_pdf_path, x_positions):
        output_pdf_path = 'saved_pdf/output_vertical.pdf'
        input_pdf = PdfReader(input_pdf_path)
        output_pdf = PdfWriter()
        for page_num in range(len(input_pdf.pages)):
            page = input_pdf.pages[page_num]
            page_width = page.mediabox.upper_right[0] - page.mediabox.lower_left[0]
            page_height = page.mediabox.upper_right[1] - page.mediabox.lower_left[1]
            content = page.extract_text()
            lines = content.split('\n')
            self.insert_vertical_lines(page, x_positions, page_height)
            output_pdf.add_page(page)
        with open(output_pdf_path, 'wb') as output_file:
            output_pdf.write(output_file)
        return output_pdf_path

    def check_date(self, df):
        if pd.to_datetime(df['Value Date'].iloc[-1], dayfirst=True) < pd.to_datetime(df['Value Date'].iloc[0],
                                                                                     dayfirst=True):
            new_df = df[::-1].reset_index(drop=True)
        else:
            new_df = df.copy()  # No reversal required
        return new_df

    def check_balance(self, df):
        df.loc[:, 'Debit'] = pd.to_numeric(df['Debit'], errors='coerce')  # Convert 'Debit' column to numeric
        df.loc[:, 'Credit'] = pd.to_numeric(df['Credit'], errors='coerce')  # Convert 'Credit' column to numeric
        df.loc[:, 'Balance'] = pd.to_numeric(df['Balance'], errors='coerce')  # Convert 'Balance' column to numeric

        prev_balance = df['Balance'].iloc[0]
        for index, row in df.iloc[1:].iterrows():
            current_balance = row['Balance']
            if row['Debit'] > 0 and prev_balance > 0:
                calculated_balance = prev_balance - row['Debit']
                if round(calculated_balance, 2) != round(current_balance, 2):
                    raise ValueError(f"Error at row {index}: Calculated balance ({calculated_balance}) "
                                     f"doesn't match current balance ({current_balance}) and Error at row DEBIT{row['Debit']} between {row['Value Date']} ")
            elif row['Credit'] > 0 and prev_balance > 0:
                calculated_balance = prev_balance + row['Credit']
                if round(calculated_balance, 2) != round(current_balance, 2):
                    raise ValueError(f"Error at row {index}: Calculated balance ({calculated_balance}) "
                                     f"doesn't match current balance ({current_balance}) and Error at row CREDIT{row['Credit']} between {row['Value Date']} ")
            prev_balance = current_balance
        return df

    def extract_the_df(self, idf):
        balance_row_index = idf[idf.apply(lambda row: 'balance' in ' '.join(row.astype(str)).lower(), axis=1)].index

        # Check if "Balance" row exists
        if not balance_row_index.empty:
            # Get the index of the "Balance" row
            balance_row_index = balance_row_index[0]
            # Create a new DataFrame from the "Balance" row till the end
            new_df = idf.iloc[balance_row_index:]
        else:
            return idf
        return new_df

    def uncontinuous(self, df):
        df = df[~df.apply(lambda row: row.astype(str).str.contains('Balance', case=False)).any(axis=1)]
        return df

    #################--------******************----------#####################

    ### IDBI BANK
    def idbi(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)

        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[0:2], axis=1)  # Removes column at position 1
        df = df.rename(
            columns={2: 'Value Date', 3: 'Description', 4: 'Cheque No', 5: 'CR/DR', 6: 'CCY', 7: 'Amount (INR)',
                     8: 'Balance'})
        ### change the columns of the df according to the standard format
        df['Credit'] = 0
        df['Debit'] = 0
        df.loc[df['CR/DR'] == 'Cr.', 'Credit'] = df.loc[df['CR/DR'] == 'Cr.', 'Amount (INR)']
        df.loc[df['CR/DR'] != 'Cr.', 'Debit'] = df.loc[df['CR/DR'] != 'Cr.', 'Amount (INR)']
        df = df.drop(['CR/DR', 'Amount (INR)'], axis=1)
        # Replace '/' with '-' in the 'Value Date' column
        df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'IDBI Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### AXIS BANK
    def axis(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        # start custom extraction
        df = df.rename(
            columns={0: 'Value Date', 1: 'Cheque No', 2: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance',
                     6: 'Init(Br)'})
        df = df[:-1]
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'Axis Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### SBI BANK
    def sbi(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
        df = df.rename(
            columns={0: 'Value Date', 2: 'Description', 3: 'Cheque No', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'SBI Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### IDFC BANK
    def idfc(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
        df = df.rename(
            columns={0: 'Value Date', 2: 'Description', 3: 'Cheque No', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'IDFC Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### PNB BANK
    def pnb(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.rename(
            columns={0: 'Value Date', 1: 'Cheque No', 2: 'Debit', 3: 'Credit', 4: 'Balance', 5: 'Description'})
        df['Value Date'] = df['Value Date'].astype(str).str.replace('/', '-', n=2)
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].replace({' Cr.': '', ' Dr.': ''}, regex=True)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'PNB Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### YES BANK
    def yes_bank(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[1:2], axis=1)  # Removes column at position 1
        # df = df.iloc[1:] # Removes the first 1 rows
        df = self.uncontinuous(df)
        df = df.rename(
            columns={0: 'Value Date', 2: 'Cheque No', 3: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'Yes Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### UNION BANK
    def union(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[[2, 3, 4]], axis=1)  # Removes column at position 3,5
        # df = df.iloc[1:] # Removes the first 1 rows
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
        idf['Bank'] = 'Union Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### KOTAK BANK
    def kotak(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            print("##############################")
            print(df_total)
            print("##############################")
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)

        # start custom extraction
        df = df.drop(df.columns[[0]], axis=1)  # Removes column at position 3,5
        # df = df.iloc[1:] # Removes the first 1 rows
        df = df.rename(columns={1: 'Value Date', 2: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
        df = df[df['Balance'] != ""]
        df = df.iloc[:-2]
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d %b %Y').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace('-', '')
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'Kotak Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ###BANK OF BARODA
    def bob(self, unlocked_pdf_path):
        x_positions = [80, 360, 460, 600, 700]
        lined_pdf_path = self.separate_lines_in_vertical_pdf(unlocked_pdf_path, x_positions)
        pdf = pdfplumber.open(lined_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        # start custom extraction
        df = df.drop(df.columns[[2]], axis=1)  # Removes column at position 2
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
        df = df.dropna(subset=['Balance'])

        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'Bank of Baroda'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ###ICICI BANK
    def icici(self, unlocked_pdf_path):
        df = pd.DataFrame()
        dfs = []
        with pdfplumber.open(unlocked_pdf_path) as pdf:
            num_pages = len(pdf.pages)
            print("Number of Pages in PDF:", num_pages)
            for i in range(num_pages):
                page = pdf.pages[i]
                table = page.extract_tables()
                if len(table) > 0:
                    for tab in table:
                        df = pd.DataFrame(tab)
                        dfs.append(df)
        df_total = pd.concat(dfs, ignore_index=True)
        new_df = self.extract_the_df(df_total)
        df = self.uncontinuous(new_df)
        # #start custom extraction
        df = df.rename(columns={1: 'Value Date', 4: 'Description', 5: 'Debit', 6: 'Credit', 7: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime(
            '%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Balance'] = df['Balance'].str.replace('-', '')
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'].notna() & (df['Description'] != '') & (df['Description'] != 'None')]
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
        idf['Bank'] = 'ICICI Bank'

        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ###Indusind BANK
    def indus(self, unlocked_pdf_path):
        df = pd.DataFrame()
        dfs = []

        with pdfplumber.open(unlocked_pdf_path) as pdf:
            num_pages = len(pdf.pages)
            print("Number of Pages in PDF:", num_pages)

            for i in range(num_pages):
                page = pdf.pages[i]
                table = page.extract_tables()

                if len(table) > 0:
                    for tab in table:
                        df = pd.DataFrame(tab)
                        dfs.append(df)
        df = pd.concat(dfs, ignore_index=True)
        new_df = self.extract_the_df(df)
        df = self.uncontinuous(new_df)
        # start custom extraction
        df = df.iloc[1:]  # Removes the first 1 rows
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
        df = df[df['Description'].notnull()]
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%b-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]
        idf['Bank'] = 'IndusInd Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ###HDFC BANK
    def hdfc(self, unlocked_pdf_path):
        lined_pdf_path = self.separate_lines_in_pdf(unlocked_pdf_path)
        pdf = pdfplumber.open(lined_pdf_path)

        df_total = pd.DataFrame()
        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
            df_total = df_total.replace('', np.nan, regex=True)
        w = df_total.drop_duplicates()
        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        df['new_column'] = np.nan
        counter = 0
        # Iterate over the dataframe rows
        for index, row in df.iterrows():
            if pd.notnull(row[0]):
                counter += 1
            df.at[index, 'new_column'] = counter
        # Iterate over the dataframe rows
        for index, row in df.iterrows():
            if pd.isna(row[0]):
                df.at[index, 'new_column'] = np.NaN
        df['new_column'].fillna(method='ffill', inplace=True)
        df[1].fillna('', inplace=True)
        df[1] = df.groupby('new_column')[1].transform(lambda x: ' '.join(x))
        df = df.drop_duplicates(subset='new_column').reset_index(drop=True)
        df = df.drop([2, 3]).reset_index(drop=True)
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 4: 'Debit', 5: 'Credit', 6: 'Balance'})
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%y', errors='coerce').dt.strftime(
            '%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df.dropna(subset=["Value Date"])

        # Reorder the columns
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'HDFC Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### NKGSB BANK
    def nkgsb(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()
        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        # start custom extraction
        df = df.iloc[:-1]  # Removes the first 1 rows
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'NKGSB Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    ### INDIAN BANK
    def indian(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        # start custom extraction
        # df = df.iloc[1:] # Removes the first 1 rows
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 2: 'Debit', 3: 'Credit', 4: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        df = df[pd.notna(df['Value Date'])]
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'Indian Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
    ### TJSB BANK
    def tjsb(self, unlocked_pdf_path):
        pdf = pdfplumber.open(unlocked_pdf_path)
        df_total = pd.DataFrame()

        for i in range(len(pdf.pages)):
            p0 = pdf.pages[i]
            table = p0.extract_table()
            df_total = df_total._append(table, ignore_index=True)
            df_total.replace({r'\n': ' '}, regex=True, inplace=True)
        w = df_total.drop_duplicates()

        new_df = self.extract_the_df(w)
        df = self.uncontinuous(new_df)
        # start custom extraction
        # df = df.iloc[1:] # Removes the first 1 rows
        df = df.rename(columns={0: 'Value Date', 1: 'Description', 3: 'Debit', 4: 'Credit', 5: 'Balance'})
        # date
        df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d-%m-%Y')
        df = self.check_date(df)
        df['Balance'] = df['Balance'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = df['Debit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Credit'] = df['Credit'].str.replace(r'[^\d.-]+', '', regex=True)
        df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
        df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')
        df['Balance'] = pd.to_numeric(df['Balance'], errors='coerce')
        df = df[df['Description'] != '']
        idf = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance']]

        idf['Bank'] = 'TJSB Bank'
        return idf

    # -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

    #################--------******************------------#####################

    # Function to extract text from a PDF file
    def extract_text_from_pdf(self, unlocked_file_path):
        with open(unlocked_file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ''
            for page in pdf_reader.pages:
                text += page.extract_text()
        return text

    def check_statement_period_monthwise(self, start_date_str, end_date_str):
        date_format = "%d-%m-%Y"
        start_date = datetime.strptime(start_date_str, date_format)
        end_date = datetime.strptime(end_date_str, date_format)
        if start_date.day != 1:
            print(
                f"The statement should start from the first day of a month. Your statement starts on {start_date_str}.")
        next_day = end_date + timedelta(days=1)
        if next_day.day != 1:
            print(f"The statement should end on the last day of a month. Your statement ends on {end_date_str}.")
        return print("Statement starts from first day of month and ends on last day of month.")

    def convert_to_dt_format(self, date_str):
        formats_to_try = ["%d-%m-%Y", "%d %b %Y", "%d %B %Y", "%d/%m/%Y", "%d/%m/%Y", "%d-%m-%Y", "%d-%b-%Y",
                          "%d/%m/%Y"]
        for format_str in formats_to_try:
            try:
                date_obj = datetime.strptime(date_str, format_str)
                return date_obj.strftime("%d-%m-%Y")
            except ValueError:
                pass
        raise ValueError("Invalid date format: {}".format(date_str))

    def find_names_and_account_numbers_sbi(self, text):
        name_pattern = re.compile(r'Name\s*:\s*([^\n]+)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account Number\s*:\s*(\d+)', re.IGNORECASE)

        pys = re.compile(r'Statement from (.*)', re.IGNORECASE)
        date_str = pys.findall(text)[0]
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)
        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_hdfc(self, text):
        # Customize the regex patterns according to your needs
        name_pattern = re.compile(r'(?:MR\.|M/S\.)\s*([^\n]+)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account No\s*:\s*(\d+)', re.IGNORECASE)

        from_index = text.find("From :") + len("From :")
        start_date = text[from_index + 1: from_index + 11].strip()
        to_index = text.find("To :") + len("To :")
        end_date = text[to_index + 1: to_index + 11].strip()
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)
        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_axis(self, text):
        joint_holder_text = "Joint Holder :"
        parts = text.split(joint_holder_text, 1)
        names = parts[0].strip().split('\n')
        account_number_pattern = re.compile(r'Account No\s*:\s*(\d+)', re.IGNORECASE)
        account_numbers = account_number_pattern.findall(text)

        date_pattern = r"\d{2}-\d{2}-\d{4}"
        dates = re.findall(date_pattern, text)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_idbi(self, text):
        name_pattern = re.compile(r': (.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'A/C NO: (\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'A/C STATUS\n(.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}/\d{2}/\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_bob(self, text):
        name_pattern = re.compile(r'Savings Account - \S+\s+(.*?)(?:\s|$)')
        account_number_pattern = re.compile(r'Account No\s*:\s*(\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'INR for the period (.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}/\d{2}/\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_icici(self, text):
        name_pattern = re.compile(r' -\s*([^\n]+)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account Number\s*(\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)
        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_union(self, text):
        name_pattern = re.compile(r'Statement of Account(.*)', re.DOTALL)
        account_number_pattern = re.compile(r'Account No (\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        lines = names[0].strip().split('\n')
        names = lines
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'Statement Period From -(.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}/\d{2}/\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_yes(self, text):
        name_pattern = re.compile(r'Primary Holder (.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'account number (\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'2. Period: (.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        start_date = datetime.strptime(py[:11], "%d %b %Y").strftime("%d %b %Y")
        end_date = datetime.strptime(py[-11:], "%d %b %Y").strftime("%d %b %Y")
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_pnb(self, text):
        name_pattern = re.compile(r'Customer Name: (.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account:(\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'Statement Period  : (.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}/\d{2}/\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_idfc(self, text):
        name_pattern = re.compile(r'CUSTOMER NAME(.*)(?=\s*:|$)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = text.split('\n')[3:]

        date_pattern = r"\d{2}-[A-Za-z]{3}-\d{4}"
        dates = re.findall(date_pattern, text)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_kotak(self, text):
        name_pattern = re.compile(r'Primary Holder (.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'account number (\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)
        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_indus(self, text):
        name_pattern = re.compile(r'  \n(.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account No                          :\n(\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'Period                                   :\n(.*)', re.IGNORECASE)
        date_str = pys.findall(text)[0]
        start_date_str = self.convert_to_dt_format(date_str.split(" To ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" To ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_tjsb(self, text):
        name_pattern = re.compile(r'Name :\n(.*)', re.IGNORECASE)
        account_number_pattern = re.compile(r'\nX(.*)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        start_date = re.compile(r'Period :\n(.*)', re.IGNORECASE).findall(text)[0]
        end_date = re.compile(r'\nTo\n(.*)', re.IGNORECASE).findall(text)[0]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_indian(self, text):
        name_pattern = re.compile(r'Customer Name  :(.*)(?=\s*CIF\s*:\s*\d+|$)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account:(\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'From (.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}/\d{2}/\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def find_names_and_account_numbers_nkgsb(self, text):
        name_pattern = re.compile(r'(?:MR\.|MRS\.)\s*([^\n]+)', re.IGNORECASE)
        account_number_pattern = re.compile(r'Account Number : (\d+)', re.IGNORECASE)
        names = name_pattern.findall(text)
        account_numbers = account_number_pattern.findall(text)

        pys = re.compile(r'Period : (.*)', re.IGNORECASE)
        py = pys.findall(text)[0]
        date_pattern = r"\d{2}-\d{2}-\d{4}"
        dates = re.findall(date_pattern, py)
        start_date, end_date = dates[0], dates[1]
        complete_date_range = f"{start_date} to {end_date}"
        date_str = complete_date_range
        start_date_str = self.convert_to_dt_format(date_str.split(" to ")[0])
        end_date_str = self.convert_to_dt_format(date_str.split(" to ")[1])
        self.check_statement_period_monthwise(start_date_str, end_date_str)

        if not names:
            names = [None]
        if not account_numbers:
            account_numbers = [None]
        listO = [names[0], account_numbers[0]]
        return listO

    def extraction_process(self, bank, pdf_path, pdf_password, start_date, end_date):
        bank = re.sub(r'\d+', '', bank)
        unlocked_pdf_path = self.unlock_the_pdfs_path(pdf_path, pdf_password)
        print(unlocked_pdf_path)
        text = self.extract_text_from_pdf(unlocked_pdf_path)

        if bank == "Axis":
            df = pd.DataFrame(self.axis(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_axis(text)

        elif bank == "IDBI":
            df = pd.DataFrame(self.idbi(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_idbi(text)

        elif bank == "SBI":
            df = pd.DataFrame(self.sbi(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_sbi(text)

        elif bank == "IDFC":
            df = pd.DataFrame(self.idfc(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_idfc(text)

        elif bank == "PNB":
            df = pd.DataFrame(self.pnb(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_pnb(text)

        elif bank == "Yes Bank":
            df = pd.DataFrame(self.yes_bank(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_yes(text)

        elif bank == "Kotak":
            df = pd.DataFrame(self.kotak(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_kotak(text)

        elif bank == "Union":
            df = pd.DataFrame(self.union(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_union(text)

        elif bank == "ICICI":
            df = pd.DataFrame(self.icici(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_icici(text)

        elif bank == "BOB":
            df = pd.DataFrame(self.bob(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_bob(text)

        elif bank == "IndusInd":
            df = pd.DataFrame(self.indus(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_indus(text)

        elif bank == "Indian":
            df = pd.DataFrame(self.indian(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_indian(text)

        elif bank == "TJSB":
            df = pd.DataFrame(self.tjsb(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_tjsb(text)

        elif bank == "NKGSB":
            df = pd.DataFrame(self.nkgsb(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_nkgsb(text)

        elif bank == "HDFC":
            df = pd.DataFrame(self.hdfc(unlocked_pdf_path))
            acc_name_n_number = self.find_names_and_account_numbers_hdfc(text)

        else:
            df = pd.NA
            acc_name_n_number = pd.NA
            raise ValueError("Bank Does not Exist")

        df = df.reset_index(drop=True)

        # if df['Value Date'].iloc[0] != start_date and df['Value Date'].iloc[-1] != end_date:
        #     print("--------@@@@@@@@@@@@@@-----------------------@@@@@@@@@@@@@@-------------")
        #     raise ValueError("The Start and End Dates provided by the user do not match ...")

        return df, acc_name_n_number

    def process_repeating_columns(self, oy):
        df = pd.concat(oy, axis=1)
        df = df.loc[:, ~df.columns.duplicated(keep='first') | (df.columns != 'Day')]
        repeating_columns = [col for col in df.columns if df.columns.tolist().count(col) > 1]

        idf = pd.DataFrame({col: df[col].sum(axis=1).round(4) for col in repeating_columns})
        df = df.drop(columns=repeating_columns)
        concatenated_df = pd.concat([df, idf], axis=1)

        sorted_columns = sorted([col for col in concatenated_df.columns if col != 'Day'],
                                key=lambda x: pd.to_datetime(x, format="%b-%Y"))
        sorted_columns_formatted = [col.strftime("%b-%Y") if isinstance(col, pd.Timestamp) else col for col in
                                    sorted_columns]
        concatenated_df = concatenated_df[['Day'] + sorted_columns_formatted]
        return concatenated_df

    def months_between(self, start_month, end_month):
        def month_sort_key(month_str):
            return datetime.strptime(month_str, '%b-%Y')

        start_ate = datetime.strptime(start_month, '%b-%Y')
        end_ate = datetime.strptime(end_month, '%b-%Y')
        months_list = []
        current_date = start_ate
        while current_date <= end_ate:
            months_list.append(current_date.strftime('%b-%Y'))
            current_date += relativedelta(months=1)
        return sorted(months_list, key=month_sort_key)

    def monthly(self, df):
        # add a new row with the average of month values in each column
        new_row = pd.DataFrame(df.iloc[0:31].mean(axis=0)).T
        monthly_avg = pd.concat([df, new_row], ignore_index=True)
        monthly_avg.iloc[-1, 0] = 'Average'
        return monthly_avg

    def eod(self, df):
        end_day = df["Date"].iloc[-1]
        df = df[["Value Date", "Balance", "Month", "Date", "Bank"]]
        bank_names = df["Bank"].unique().tolist()
        multiple_eods = []

        for bank in bank_names:
            idf = df[df["Bank"] == bank]
            result_eod = pd.DataFrame()
            for month in idf['Month'].unique():
                eod_month_df = idf.loc[idf['Month'] == month].drop_duplicates(subset='Date', keep='last')

                # Loop through each day in the month
                for day in range(1, 32):
                    # Check if there are any rows with the current day
                    day_present = False
                    for index, row in eod_month_df.iterrows():
                        if row['Date'] == day:
                            day_present = True
                            break

                    # If day is not present, add a row with NaN values for all columns except the date
                    if not day_present:
                        new_row = {'Balance': 0, 'Month': eod_month_df.iloc[0]['Month'], 'Date': day}
                        eod_month_df = pd.concat([eod_month_df, pd.DataFrame(new_row, index=[0])], ignore_index=True)
                        eod_month_df = eod_month_df.sort_values(by='Date')

                result_eod = pd.concat([result_eod, eod_month_df], ignore_index=True)

            # iterate through column and replace zeros with previous value
            previous_eod = 0
            for i, value in enumerate(result_eod['Balance']):
                if value == 0:
                    result_eod.loc[i, 'Balance'] = previous_eod
                else:
                    previous_eod = value

            pivot_df = result_eod.pivot(index='Date', columns='Month', values='Balance').reset_index(drop=True)
            column_order = idf["Month"].unique()  # do not change
            pivot_df = pivot_df.reindex(columns=column_order)
            pivot_df.insert(0, 'Day', range(1, 32))

            columns = pivot_df.columns[1:]
            col_values = ['Feb', 'Apr', 'Jun', 'Sep',
                          'Nov']  # no hard code now :: these are the months in every year not having 31 days

            for i, row in pivot_df.iterrows():
                for month in columns:
                    if any(col in month for col in col_values):
                        if 'Feb' in month and row['Day'] > 28:
                            pivot_df.loc[i, month] = np.nan
                        elif row['Day'] > 30:
                            pivot_df.loc[i, month] = np.nan

            # last_column_list = pivot_df.iloc[:, -1].tolist()
            # new_column = last_column_list.copy()
            # new_column[end_day + 1:] = [0] * (len(new_column) - end_day - 1)
            # pivot_df.iloc[:, -1] = new_column

            multiple_eods.append(pivot_df)

            if len(multiple_eods) < 1:
                adf = multiple_eods[0]
                # add a new row with the sum of values in each column
                new_row = pd.DataFrame(adf.iloc[0:31].sum(axis=0)).T
                total_df = pd.concat([adf, new_row], ignore_index=True)
                total_df.iloc[-1, 0] = 'Total'
                all_df = self.monthly(total_df)
            else:
                adf = self.process_repeating_columns(multiple_eods)
                # add a new row with the sum of values in each column
                new_row = pd.DataFrame(adf.iloc[0:31].sum(axis=0)).T
                total_df = pd.concat([adf, new_row], ignore_index=True)
                total_df.iloc[-1, 0] = 'Total'
                all_df = self.monthly(total_df)

        return all_df

    def category_add(self, df):
        # df2 = pd.read_excel('findaddy/banks/axis_category.xlsx')
        df2 = pd.read_excel("multiple_category.xlsx")

        category = []
        for desc in df['Description'].str.lower():
            match_found = False
            for value in df2["Particulars"].str.lower():
                if value in desc:
                    summary = df2.loc[df2["Particulars"].str.lower() == value, "Category"].iloc[0]
                    category.append(summary)
                    match_found = True
                    break
            if not match_found:
                category.append("Suspense")
        df['Category'] = category
        # Reorder the columns
        # df = df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']]
        return df

    ##SHEETS
    # for investment
    def total_investment(self, df):
        invest_df = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "Investment":
                invest_df = invest_df._append(row, ignore_index=True)
        return invest_df

    # for return of investment
    def redemption_investment(self, df):
        red_df = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "Redemption of Investment":
                red_df = red_df._append(row, ignore_index=True)
        return red_df

    # for cash withdrawal
    def cash_withdraw(self, df):
        cashw_df = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "Cash Withdrawal":
                cashw_df = cashw_df._append(row, ignore_index=True)
        return cashw_df

    # for cash deposit
    def cash_depo(self, df):
        cashd_df = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "Cash Deposits":
                cashd_df = cashd_df._append(row, ignore_index=True)
        return cashd_df

    # dividend/interest
    def div_int(self, df):
        iii = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "Dividend/interest":
                iii = iii._append(row, ignore_index=True)
        return iii

    # recurring emi
    def emi(self, df):
        em_i = pd.DataFrame()
        for index, row in df.iterrows():
            arow = row["Category"]
            if arow == "EMI":
                em_i = em_i._append(row, ignore_index=True)
        return em_i

    # for creditor list

    #
    # for suspense credit
    def suspense_credit(self, df):
        suspense_cr = df[df["Category"].str.contains("Suspense")].groupby('Credit')
        suspense_cr = suspense_cr.apply(lambda x: x)
        return suspense_cr
        # c_df = pd.DataFrame()
        # for index, row in df.iterrows():
        #     credit_amount = pd.to_numeric(row['Credit'], errors='coerce')
        #     arow = row["Category"]
        #     if arow == "Suspense" and credit_amount > 0:
        #         c_df = c_df._append(row, ignore_index=True)
        # return c_df

    # for suspense debit

    def suspense_debit(self, df):
        suspense_dr = df[df["Category"].str.contains("Suspense")].groupby('Debit')
        suspense_dr = suspense_dr.apply(lambda x: x)
        return suspense_dr
        # d_df = pd.DataFrame()
        # for index, row in df.iterrows():
        #     debit_amount = pd.to_numeric(row['Debit'], errors='coerce')
        #     arow = row["Category"]
        #     if arow == "Suspense" and debit_amount > 0:
        #         d_df = d_df._append(row, ignore_index=True)
        # return d_df

    # ***************/-first page summary sheet-/*********************#
    def avgs_df(self, df):
        # quarterly_avg
        if df.shape[1] > 3:
            df_chi_list_1 = []
            # Iterate through every three columns in the original DataFrame
            for i in range(1, df.shape[1], 3):
                # Get the current three columns
                subset = df.iloc[:, i:i + 3]
                if subset.shape[1] < 3:
                    new_row = 0.0
                else:
                    new_row = subset.iloc[-2].sum() / 3
                subset.loc[len(subset)] = new_row
                df_chi_list_1.append(subset)
            result_df = pd.concat(df_chi_list_1, axis=1)
            new_row = pd.Series({'Day': 'Quarterly Average'})
            df = df._append(new_row, ignore_index=True)
            result_df.insert(0, 'Day', df['Day'])
            df = result_df

            # half - yearly avg
            if df.shape[1] > 6:
                df_chi_list_2 = []
                # Iterate through every three columns in the original DataFrame
                for i in range(1, df.shape[1], 6):
                    # Get the current three columns
                    subset = df.iloc[:, i:i + 6]
                    if subset.shape[1] < 6:
                        new_row = 0.0
                    else:
                        new_row = subset.iloc[-3].sum() / 6
                    subset.loc[len(subset)] = new_row
                    df_chi_list_2.append(subset)
                result_df = pd.concat(df_chi_list_2, axis=1)
                new_row = pd.Series({'Day': 'Half-Yearly Average'})
                df = df._append(new_row, ignore_index=True)
                result_df.insert(0, 'Day', df['Day'])
                df = result_df

                # yearly avg
                if df.shape[1] > 12:
                    df_chi_list_3 = []
                    # Iterate through every three columns in the original DataFrame
                    for i in range(1, df.shape[1], 12):
                        # Get the current three columns
                        subset = df.iloc[:, i:i + 12]
                        if subset.shape[1] < 12:
                            new_row = 0.0
                        else:
                            new_row = subset.iloc[-4].sum() / 12
                        subset.loc[len(subset)] = new_row
                        df_chi_list_3.append(subset)
                    result_df = pd.concat(df_chi_list_3, axis=1)
                    new_row = pd.Series({'Day': 'Yearly Average'})
                    df = df._append(new_row, ignore_index=True)
                    result_df.insert(0, 'Day', df['Day'])
                    df = result_df


                else:
                    new_row = pd.Series({'Day': 'Yearly Average'})
                    df = df._append(new_row, ignore_index=True)

            else:
                new_row = pd.Series({'Day': 'Half-Yearly Average'})
                df = df._append(new_row, ignore_index=True)

        else:
            new_row = pd.Series({'Day': 'Quarterly Average'})
            df = df._append(new_row, ignore_index=True)

        return df

    def summary_sheet(self, idf, open_bal, close_bal, eod):
        opening_bal = open_bal
        closing_bal = close_bal
        eod_avg_df = self.avgs_df(eod) ##eod added with quarterly, half-yearly, yearly averages


        def total_number_cr(df):
            number = df["Credit"].count()
            return number

        # total amount of credit transactions
        def total_amount_cr(df):
            sum = df["Credit"].sum(axis=0)
            return sum

        def total_number_dr(df):
            number = df["Debit"].count()
            return number

        # total amount of debit transactions
        def total_amount_dr(df):
            sum = df["Debit"].sum(axis=0)
            return sum

        def total_number_cd(df):
            cd = df["Category"] == "Cash Deposits"
            cd = cd.count()
            return cd

        # total amount of cash deposits ###money credited to your account
        def total_amount_cd(df):
            amount = 0
            for index, row in df.iterrows():
                credit_amount = row['Credit']
                if row["Category"] == "Cash Deposits" and credit_amount > 0:
                    amount += credit_amount
            return amount

        def total_number_cw(df):
            cw = df["Category"] == "Cash Withdrawal"
            cw = cw.count()
            return cw

        # total amount of cash withdrawn ### money is debited from your account
        def total_amount_cw(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "Cash Withdrawal" and debit_amount > 0:
                    amount += debit_amount
            return amount
        def no_cheque_depos(df):
            return 0

        def amt_cheque_depos(df):
            return 0
        def no_cash_issued(df):

            return 0

        def total_amount_cash_issued(df):
            return 0

        def inward_cheque_bounces(df):
            return 0

        def outward_cheque_bounces(df):
            return 0

        def min_eod(df, month):
            eod_df = eod.iloc[:-2]
            eod_month = eod_df[month].values
            min = np.nanmin(eod_month)
            return min

        def max_eod(df, month):
            eod_df = eod.iloc[:-2]
            eod_month = eod_df[month].values
            max = np.nanmax(eod_month)
            return max

        def avg_eod(df, month):
            eod_df = eod.iloc[:-2]
            eod_month = eod_df[month].values
            avg = np.nanmean(eod_month)
            return avg

        def qtrly_avg_bal(df, month):
            qtrly_avg_balance = eod_avg_df.loc[eod_avg_df.index[-3], month]
            return qtrly_avg_balance

        def half_yrly_bal(df, month):
            half_yrly_avg_balance = eod_avg_df.loc[eod_avg_df.index[-2], month]
            return half_yrly_avg_balance

        def yrly_avg_bal(df, month):
            yrly_avg_balance = eod_avg_df.loc[eod_avg_df.index[-1], month]
            return yrly_avg_balance

        def all_bank_avg(df):
            return 0

        def top5_funds_rec(df):
            return 0

        def top5_redemption(df):
            return 0

        def bounced_txns(df):
            return 0

        def emi1(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "EMI" and debit_amount > 0:
                    amount += debit_amount
            return amount

        def total_amount_pos_cr(df):
            amount = 0
            for index, row in df.iterrows():
                credit_amount = row['Credit']
                if row["Category"] == "POS-cr" and credit_amount > 0:
                    amount += credit_amount
            return amount

        # POS transaction dr ### money is debited from your account
        def total_amount_pos_dr(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "POS-dr" and debit_amount > 0:
                    amount += debit_amount
            return amount

        def datewise_avg_bal(df):
            return 0

        # investment (money debited in total)
        def total_investment(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "Investment" and debit_amount > 0:
                    amount += debit_amount
            return amount

        # interest recieved fropm bank
        def received_interest(df):
            amount = 0
            for index, row in df.iterrows():
                credit_amount = row['Credit']
                if row["Category"] == "Interest Credit" and credit_amount > 0:
                    amount += credit_amount
            return amount

        # salary recieved
        def received_salary(df):
            amount = 0
            for index, row in df.iterrows():
                credit_amount = row['Credit']
                if row["Category"] == "Salary Received" and credit_amount > 0:
                    amount += credit_amount
            return amount

        def diff_bank_abb(df):
            return 0

        def interest_rece(df):
            return 0

        def paid_interest1(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "Interest Debit" and debit_amount > 0:
                    amount += debit_amount
            return amount

            # salary paid

        def paid_salary(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "Salary paid" and debit_amount > 0:
                    amount += debit_amount
            return amount

        def salary_rec(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Credit']
                if row["Category"] == "Salary" and debit_amount > 0:
                    amount += debit_amount
            return amount

        def paid_tds1(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "TDS" and debit_amount > 0:
                    amount += debit_amount
            return amount

        def GST(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "GST Paid" and debit_amount > 0:
                    amount += debit_amount
            return amount

        # suspense
        def suspenses(df):
            amount = 0
            for index, row in df.iterrows():
                debit_amount = row['Debit']
                if row["Category"] == "Suspense" and debit_amount > 0:
                    amount += debit_amount
            return amount

        # -----------------------X----------------------------------#
        idf['Credit'] = pd.to_numeric(idf['Credit'], errors='coerce')
        idf['Debit'] = pd.to_numeric(idf['Debit'], errors='coerce')
        # idf['Month'] = idf['Value Date'].dt.strftime('%b-%Y')
        months = idf["Month"].unique()

        number_cr = {}
        amount_cr = {}
        number_dr = {}
        amount_dr = {}
        number_cd = {}
        amount_cd = {}
        number_cw = {}
        amount_cw = {}
        no_cheque_depo ={}
        amt_cheque_depo ={}

        number_cash_issued = {}
        amount_cash_issued = {}

        inward_bounce = {}
        outward_bounce = {}
        min_eod_bal = {}
        max_eod_bal = {}
        avg_eod_bal = {}
        qtrlu_bal = {}
        half_bal = {}
        yrly_bal = {}
        all_bank_avg_bal = {}
        top_5_funds = {}
        top_5_redemptions = {}
        bounced = {}
        emi = {}
        amount_pos_cr = {}
        amount_pos_dr = {}
        datewise_bal = {}
        investment_dr = {}
        diff_bank_ab = {}
        interest_rec = {}
        paid_interest = {}
        paid_sal = {}
        received_salary = {}
        paid_tds = {}
        paid_gst = {}
        investment_cr = {}
        suspense = {}

        for month in months:
            new_df = idf[idf['Month'] == month].reset_index(drop=True)
            number_cr.update({month: total_number_cr(new_df)})
            amount_cr.update({month: total_amount_cr(new_df)})
            number_dr.update({month: total_number_dr(new_df)})
            amount_dr.update({month: total_amount_dr(new_df)})
            number_cd.update({month: total_number_cd(new_df)})
            amount_cd.update({month: total_amount_cd(new_df)})
            number_cw.update({month: total_number_cw(new_df)})
            amount_cw.update({month: total_amount_cw(new_df)})
            no_cheque_depo.update({month: no_cheque_depos(new_df)})
            amt_cheque_depo.update({month: amt_cheque_depos(new_df)})
            number_cash_issued.update({month: no_cash_issued(new_df)})
            amount_cash_issued.update({month: total_amount_cash_issued(new_df)})
            inward_bounce.update({month: inward_cheque_bounces(new_df)})
            outward_bounce.update({month: outward_cheque_bounces(new_df)})
            min_eod_bal.update({month: min_eod(new_df, month)})
            max_eod_bal.update({month: max_eod(new_df, month)})
            avg_eod_bal.update({month: avg_eod(new_df, month)})
            qtrlu_bal.update({month: qtrly_avg_bal(new_df, month)})
            half_bal.update({month: half_yrly_bal(new_df, month)})
            yrly_bal.update({month: yrly_avg_bal(new_df, month)})
            all_bank_avg_bal.update({month: all_bank_avg(new_df)})
            top_5_funds.update({month: top5_funds_rec(new_df)})
            top_5_redemptions.update({month: top5_redemption(new_df)})
            bounced.update({month: bounced_txns(new_df)})
            emi.update({month: emi1(new_df)})
            amount_pos_cr.update({month: total_amount_pos_cr(new_df)})
            amount_pos_dr.update({month: total_amount_pos_dr(new_df)})
            datewise_bal.update({month: datewise_avg_bal(new_df)})
            investment_dr.update({month: total_investment(new_df)})
            diff_bank_ab.update({month: diff_bank_abb(new_df)})
            interest_rec.update({month: interest_rece(new_df)})
            paid_interest.update({month: paid_interest1(new_df)})
            paid_sal.update({month: paid_salary(new_df)})
            received_salary.update({month: salary_rec(new_df)})
            paid_tds.update({month: paid_tds1(new_df)})
            paid_gst.update({month: GST(new_df)})



            ###now we make sheets
            sheet_1 = pd.DataFrame(
                [number_cr, amount_cr,
                 number_dr, amount_dr,
                 number_cd, amount_cd,
                 number_cw, amount_cw,
                 no_cheque_depo,amt_cheque_depo,
                 number_cash_issued, amount_cash_issued,
                 inward_bounce, outward_bounce,
                 min_eod_bal,max_eod_bal,
                 avg_eod_bal,qtrlu_bal,
                 half_bal,yrly_bal,
                 all_bank_avg_bal,top_5_funds,
                 top_5_redemptions,bounced,
                 emi,amount_pos_cr,amount_pos_dr,
                 datewise_bal,investment_dr,
                 diff_bank_ab,interest_rec,
                 paid_interest,paid_sal,
                 received_salary,paid_tds,
                 opening_bal,closing_bal,paid_gst])
            sheet_1.insert(0, "Particulars",
                           ["Total No. of Credit Transactions","Total Amount of Credit Transactions",
                            "Total No. of Debit Transactions","Total Amount of Debit Transactions",
                            "Total No. of Cash Deposits","Total Amount of Cash Deposits",
                            "Total No. of Cash Withdrawals","Total Amount of Cash Withdrawals",
                            "Total No. of Cheque Deposits","Total Amount of Cheque Deposits",
                            "Total No. of Cheque Issued","Total Amount of Cheque Issued",
                            "Total No. of Inward Cheque Bounces","Total No. of Outward Cheque Bounces",
                            "Min EOD Balance","Max EOD Balance",
                            "Average EOD Balance","Qtrly AVG Bal",
                            "Half Yrly AVG Bal","Yrly AVG Bal",
                            "All Bank Avg Balance","Monthly Top 5 Funds Received",
                            "Monthly Top 5 Funds Remittances","Bounced Txns",
                            "EMI","Pos-cr","Pos-dr",
                            "Datewise Average Balance","Investment Details",
                            "Different Bank ABB Balance","Bank Interest Received",
                            "Bank Interest Paid (Only in OD/CC A/c)","Salaries Paid",
                            "Salary Received","TDS Deducted",
                            "Opening Balance","Closing Balance",
                            "Total GST"
                            ])
            sheet_1['Total'] = sheet_1.iloc[:, 1:].sum(axis=1)

            df_list = [sheet_1]

        return df_list

    def process_transaction_sheet_df(self, df):
        start_month = df["Month"].iloc[0]
        end_month = df["Month"].iloc[-1]
        A = pd.date_range(start=start_month, end=end_month, freq='M').strftime('%b-%Y')
        B = df["Month"].tolist()
        results = list(set(A) - set(B))
        new_data = {
            'Value Date': [0] * len(results),
            'Description': ["None"] * len(results),
            'Debit': [0] * len(results),
            'Credit': [0] * len(results),
            'Balance': [0] * len(results),
            'Month': results,
            'Date': [1] * len(results),
            'Category': ["None"] * len(results),
            'Bank': [0] * len(results)
        }
        odf = pd.DataFrame(new_data)
        idf = pd.concat([df, odf], ignore_index=True)
        idf["Month"] = pd.to_datetime(idf["Month"], format="%b-%Y")
        adf = idf.copy()
        idf.sort_values(by="Month", inplace=True)
        idf["Month"] = idf["Month"].dt.strftime("%b-%Y")
        idf.reset_index(drop=True, inplace=True)
        for index, row in idf.iterrows():
            if row['Bank'] == 0 and row['Balance'] == 0:
                idf.at[index, 'Bank'] = idf.at[index - 1, 'Bank']
                idf.at[index, 'Balance'] = idf.at[index - 1, 'Balance']
        B = idf.copy()
        none_rows_positions = B[B['Category'] == 'None'].index.tolist()
        rows_as_dict = {}
        for position in none_rows_positions:
            row_as_dict = B.loc[position].to_dict()
            rows_as_dict[position] = row_as_dict
        for index, row_data in rows_as_dict.items():
            if index in adf.index:
                adf.loc[index + 1:] = adf.loc[index:].shift(1)
            # Insert the new row at the given index
            adf.loc[index] = row_data
        adf.reset_index(drop=True, inplace=True)
        idf = adf.copy()
        idf["Month"] = idf["Month"].dt.strftime("%b-%Y")
        idf.reset_index(drop=True, inplace=True)
        for index, row in idf.iterrows():
            if row['Category'] == "None":
                idf.at[index, 'Bank'] = idf.at[index - 1, 'Bank']
                idf.at[index, 'Balance'] = idf.at[index - 1, 'Balance']
        return idf

    def Single_Bank_statement(self, dfs, name_dfs):
        data = []
        for key, value in name_dfs.items():
            bank_name = key
            acc_name = value[0]
            acc_num = value[1]
            if acc_num == "None":
                masked_acc_num = "None"
            else:
                masked_acc_num = "X" * (len(acc_num) - 4) + acc_num[-4:]
            data.append([masked_acc_num, acc_name, bank_name])

        name_n_num_df = pd.DataFrame(data, columns=['Account Number', 'Account Name', 'Bank'])
        num_pairs = len(pd.Series(dfs).to_dict())

        # print(dfs.values())

        concatenated_df = pd.concat(list(dfs.values()))
        concatenated_df = concatenated_df.fillna('')
        concatenated_df['Value Date'] = pd.to_datetime(concatenated_df['Value Date'], format='%d-%m-%Y',
                                                       errors='coerce')
        concatenated_df['Month'] = concatenated_df['Value Date'].dt.strftime('%b-%Y')
        concatenated_df['Date'] = concatenated_df['Value Date'].dt.day
        # df = concatenated_df.sort_values(by='Value Date',  ascending=True).reset_index(drop=True)
        concatenated_df.drop_duplicates(keep='first', inplace=True)
        df = concatenated_df.reset_index(drop=True)

        old_transaction_sheet_df = self.category_add(df)
        transaction_sheet_df = self.process_transaction_sheet_df(old_transaction_sheet_df)
        old_excel_transaction_sheet_df = old_transaction_sheet_df[
            ['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']]

        excel_transaction_sheet_df = self.check_balance(old_excel_transaction_sheet_df)

        eod_sheet_df = self.eod(transaction_sheet_df)
        # #opening & closing balance
        opening_bal = eod_sheet_df.iloc[0, 1:].to_dict()
        closing_bal = {}
        for column in eod_sheet_df.columns[1:]:
            non_zero_rows = eod_sheet_df.loc[eod_sheet_df[column] != 0]
            if len(non_zero_rows) > 0:
                last_non_zero_row = non_zero_rows.iloc[-1]
                closing_bal[column] = last_non_zero_row[column]
        # for summary sheets
        summary_df_list = self.summary_sheet(transaction_sheet_df.copy(), opening_bal, closing_bal, eod_sheet_df.copy())
        sheet_name = "Summary"  # summary joining
        name_n_num_df.to_excel(self.writer, sheet_name=sheet_name, startcol=1, index=False)
        summary_df_list[0].to_excel(self.writer, sheet_name=sheet_name, startrow=name_n_num_df.shape[0] + 2, startcol=1,
                                    index=False)

        if num_pairs > 1:
            excel_transaction_sheet_df.to_excel(self.writer, sheet_name='Combined Transaction', index=False)
            eod_sheet_df.to_excel(self.writer, sheet_name='Combined EOD Balance', index=False)
        else:
            sheets_oNc_list = []
            for key, value in dfs.items():
                bank_name = key
                df = pd.DataFrame(value)
                df = df.fillna('')
                # Convert 'Value Date' column to datetime format
                df['Value Date'] = pd.to_datetime(df['Value Date'], format='%d-%m-%Y', errors='coerce')
                df['Month'] = df['Value Date'].dt.strftime('%b-%Y')
                df['Date'] = df['Value Date'].dt.day
                df = df.reset_index(drop=True)
                old_transaction_sheet_df = self.category_add(df)
                transaction_sheet_df = self.process_transaction_sheet_df(old_transaction_sheet_df)
                excel_transaction_sheet_df = old_transaction_sheet_df[
                    ['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']]
                excel_transaction_sheet_df.to_excel(self.writer, sheet_name=f'{bank_name} Transaction', index=False)
                eod_sheet_df = self.eod(transaction_sheet_df)
                eod_sheet_df.to_excel(self.writer, sheet_name=f'{bank_name} EOD Balance', index=False)
                # #opening & closing balance
                eod_sheet_df_2 = eod_sheet_df.iloc[:-2]
                opening_bal = eod_sheet_df_2.iloc[0, 1:].to_dict()
                closing_bal = {}
                for column in eod_sheet_df_2.columns[1:]:
                    non_zero_rows = eod_sheet_df_2.loc[eod_sheet_df_2[column] != 0]
                    if len(non_zero_rows) > 0:
                        last_non_zero_row = non_zero_rows.iloc[-1]
                        closing_bal[column] = last_non_zero_row[column]
                sheet_1 = pd.DataFrame([opening_bal, closing_bal])
                sheet_1.insert(0, bank_name, ["Opening Balance", "Closing Balance"])
                sheet_1['Total'] = sheet_1.iloc[:, 1:].sum(axis=1)
                sheets_oNc_list.append(sheet_1)
                # sheet_name = "Opening and Closing Balance"  # summary joining
                # start_row = 0
                # for sheet in sheets_oNc_list:
                #     sheet.to_excel(self.writer, sheet_name=sheet_name, startrow=start_row, index=False)
                #     start_row += sheet.shape[0] + 2

        investment_df = self.total_investment(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        investment_df.to_excel(self.writer, sheet_name='Investment', index=False)
        redemption_investment_df = self.redemption_investment(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        redemption_investment_df.to_excel(self.writer, sheet_name='Redemption of Investment', index=False)
        # creditor_df = self.creditor_list(
        #     transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        # creditor_df.to_excel(self.writer, sheet_name='Creditor List', index=False)
        # debtor_df = self.debtor_list(
        #     transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        # debtor_df.to_excel(self.writer, sheet_name='Debtor List', index=False)False
        cash_withdrawal_df = self.cash_withdraw(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        cash_withdrawal_df.to_excel(self.writer, sheet_name='Cash Withdrawal', index=False)
        cash_deposit_df = self.cash_depo(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        cash_deposit_df.to_excel(self.writer, sheet_name='Cash Deposit', index=False)
        dividend_int_df = self.div_int(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        dividend_int_df.to_excel(self.writer, sheet_name='Dividend-Interest', index=False)
        emi_df = self.emi(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        emi_df.to_excel(self.writer, sheet_name='Recurring EMI', index=False)
        suspense_credit_df = self.suspense_credit(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        suspense_credit_df.to_excel(self.writer, sheet_name='Suspense Credit', index=False)
        suspense_debit_df = self.suspense_debit(
            transaction_sheet_df[['Value Date', 'Description', 'Debit', 'Credit', 'Balance', 'Category', 'Bank']])
        suspense_debit_df.to_excel(self.writer, sheet_name='Suspense Debit', index=False)

    def start_extraction(self):
        dfs = {}
        name_dfs = {}
        i = 0
        for bank in self.bank_names:
            bank = str(f"{bank}{i}")
            pdf_path = self.pdf_paths[i]
            pdf_password = self.pdf_passwords[i]
            start_date = self.start_date[i]
            end_date = self.end_date[i]
            dfs[bank], name_dfs[bank] = self.extraction_process(bank, pdf_path, pdf_password, start_date, end_date)
            i += 1
        print('|------------------------------|')
        print(self.account_number)
        print('|------------------------------|')
        # file_name = os.path.join('Excel_Files', f'BankStatement_{self.account_number}.xlsx')
        file_name = "saved_excel/Single_Extracted_statements_file_new.xlsx"
        self.writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

        self.Single_Bank_statement(dfs, name_dfs)
        self.writer._save()


bank_names = ["Axis"]
pdf_paths = ["bank_pdfs/sbi_month_wise/axis1.pdf"]
passwords = [""]
# dates should be in the format dd-mm-yy
start_date = [""]
end_date = [""]
converter = SingleBankStatementConverter(bank_names, pdf_paths, passwords, start_date, end_date, '00000037039495417',
                                         'test.py')
converter.start_extraction()
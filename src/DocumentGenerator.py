from abc import ABC, abstractmethod
import pandas as pd
import re
from docx import Document
from datetime import datetime

class DocumentGeneratorClass(ABC):
    """
    Abstract base class for generating Word documents
    from strategic Excel data.
    """

    def __init__(self, excel_path, template_path):
        self.excel_path = excel_path
        self.template_path = template_path
        self.df = None
        self.columns = None

    @abstractmethod
    def read_excel(self):
        """
        Read the Excel file and return a cleaned dataframe.
        Override this for custom Excel layouts.
        """
        pass

    @abstractmethod
    def extract_activity_columns(self):
        """
        Map columns to target/completed fields.
        Override this if column patterns differ.
        """
        pass

    @abstractmethod
    def populate_table(self, doc, row_data):
        """
        Populate the Word template table with data from a row.
        Override this if table structure differs.
        """
        pass

    @abstractmethod
    def update_paragraphs(self, doc, row_data, date_obj):
        """
        Update static placeholders in the Word document.
        Override for different template text patterns.
        """
        pass

    @abstractmethod
    def get_date_from_columns(self):
        """
        Extract date from activity/completed columns.
        Override if the date is in a different format or location.
        Should return a string formatted as '%d-%m-%Y' or '%d-%m-%y'.
        """
        pass

    def clean_name(self, text):
        text = str(text).strip().lower()
        text = text.replace("%", "percentage")
        text = re.sub(r"[^\w\s]", "", text)
        text = re.sub(r"\s+", "_", text)
        return text

    def clean_excel_rows(self, df):
        """
        Remove rows like 'Totals' or others.
        Can be overridden if needed.
        """
        return df[df[df.columns[0]] != "Totals"]

    def generate_documents(self):
        """
        Main driver method. Reads Excel, iterates rows, creates documents.
        """
        self.df = self.read_excel()
        self.df = self.clean_excel_rows(self.df)
        self.columns = self.df.columns
        print()

        # Get the date (delegated to subclass)
        date_obj = self.get_date_from_columns()

        doc_dict = {}
        for _, row in self.df.iterrows():
            division = row.iloc[0]
            doc = Document(self.template_path)
            self.update_paragraphs(doc, row, date_obj)
            self.populate_table(doc, row)
            doc_dict[division] = doc

        return doc_dict
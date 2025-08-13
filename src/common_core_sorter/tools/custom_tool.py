"""
Google Sheets and Docs Tools for CrewAI

This module provides simplified tools for:
- Writing data to Google Sheets
- Reading content from Google Docs via URL

Requirements:
- Google service account credentials
- Environment variables for authentication

Author: Ian
Last Modified: 2025-01-28
"""

from crewai.tools import BaseTool
from pydantic import BaseModel, Field
from typing import List, Type
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from dotenv import load_dotenv
import os
import re
from io import BytesIO


class GoogleSheetsWriterInput(BaseModel):
    """Schema for Google Sheets write operation parameters"""
    spreadsheet_id: str = Field(description="The ID of the Google Spreadsheet")
    column_a_value: str = Field(description="The value in column A to identify the row")
    column_letter: str = Field(description="The column letter to write to (e.g., 'P')")
    value: str = Field(description="The value to write")


class GoogleSheetsWriterTool(BaseTool):
    """
    Tool for writing data to Google Sheets.
    
    Provides functionality to append data to a specified Google Sheet.
    
    Environment Variables Required:
        GOOGLE_SHEETS_CREDENTIALS: Path to the service account credentials JSON file
    
    Usage:
        tool = GoogleSheetsWriterTool()
        result = tool.run(
            spreadsheet_id="your_spreadsheet_id",
            range="Sheet1!A1",
            values=[["data1", "data2"], ["data3", "data4"]]
        )
    """
    name: str = "Google Sheets Writer"
    description: str = "Writes data to a Google Sheets spreadsheet"
    args_schema: Type[BaseModel] = GoogleSheetsWriterInput

    def _initialize_sheets_service(self):
        """Initialize and return a Google Sheets API service instance"""
        load_dotenv()
        credentials_path = os.getenv("GOOGLE_SHEETS_CREDENTIALS")
        if not credentials_path:
            raise ValueError("GOOGLE_SHEETS_CREDENTIALS environment variable is not set")
        
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        return build('sheets', 'v4', credentials=credentials)

    def _run(self, spreadsheet_id: str, column_a_value: str, column_letter: str, value: str) -> str:
        """
        Write data to a specific cell by finding the row with matching column A value
        
        Args:
            spreadsheet_id: The ID of the Google Spreadsheet
            column_a_value: The value to search for in column A
            column_letter: The column letter to write to (e.g., 'P')
            value: The value to write
            
        Returns:
            Status message
        """
        try:
            service = self._initialize_sheets_service()
            sheet = service.spreadsheets()
            
            # First, get all values in column A to find the correct row
            tab_name = os.getenv('SHEETS_TAB_NAME', 'Sheet1')
            # Quote the tab name in case it contains spaces/special characters
            quoted_tab = f"'{tab_name}'"
            result = sheet.values().get(
                spreadsheetId=spreadsheet_id,
                range=f"{quoted_tab}!A:A"
            ).execute()
            
            column_a_values = result.get('values', [])
            
            # Find the row number where column A matches our value
            row_number = None
            for idx, row in enumerate(column_a_values):
                if row and str(row[0]) == str(column_a_value):
                    row_number = idx + 1  # Sheets are 1-indexed
                    break
            
            if row_number is None:
                return f"Error: Could not find row with column A value '{column_a_value}'"
            
            # Construct the range for the target cell
            target_range = f"{quoted_tab}!{column_letter}{row_number}"
            
            # Write the value to the target cell
            body = {
                'values': [[value]]
            }
            
            sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=target_range,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            return f"Successfully wrote to cell {target_range} (row where column A = '{column_a_value}')"
        
        except Exception as e:
            return f"Error writing to Google Sheets: {str(e)}"


class GoogleDocReaderInput(BaseModel):
    """Schema for Google Doc reader parameters"""
    doc_url: str = Field(description="The URL of the Google Doc to read")


class GoogleDocReaderTool(BaseTool):
    """
    Tool for reading content from Google Docs via URL.
    
    Extracts and returns the text content from a Google Doc.
    
    Environment Variables Required:
        GOOGLE_SHEETS_CREDENTIALS: Path to the service account credentials JSON file
        (Also works for Google Docs API)
    
    Usage:
        tool = GoogleDocReaderTool()
        content = tool.run(doc_url="https://docs.google.com/document/d/...")
    """
    name: str = "Google Doc Reader"
    description: str = "Reads and returns the content of a Google Doc from its URL"
    args_schema: Type[BaseModel] = GoogleDocReaderInput

    def _extract_doc_id(self, url: str) -> str:
        """Extract the document ID from a Google Docs URL"""
        # Pattern for Google Docs URLs
        patterns = [
            r'/document/d/([a-zA-Z0-9-_]+)',
            r'/d/([a-zA-Z0-9-_]+)',
            r'id=([a-zA-Z0-9-_]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        
        raise ValueError(f"Could not extract document ID from URL: {url}")

    def _initialize_docs_service(self):
        """Initialize and return a Google Docs API service instance"""
        load_dotenv()
        credentials_path = os.getenv("GOOGLE_SHEETS_CREDENTIALS")
        if not credentials_path:
            raise ValueError("GOOGLE_SHEETS_CREDENTIALS environment variable is not set")
        
        credentials = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=[
                'https://www.googleapis.com/auth/documents.readonly',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )
        return build('docs', 'v1', credentials=credentials)

    def _extract_text_from_elements(self, elements) -> str:
        texts = []
        for element in elements:
            if 'paragraph' in element:
                for text_element in element['paragraph'].get('elements', []):
                    tr = text_element.get('textRun')
                    if tr and 'content' in tr:
                        texts.append(tr['content'])
            elif 'table' in element:
                table = element['table']
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        texts.append(self._extract_text_from_elements(cell.get('content', [])))
            elif 'tableOfContents' in element:
                toc = element['tableOfContents']
                texts.append(self._extract_text_from_elements(toc.get('content', [])))
        return ''.join(texts)

    def _run(self, doc_url: str) -> str:
        """
        Read content from a Google Doc
        
        Args:
            doc_url: The URL of the Google Doc
            
        Returns:
            The text content of the document
        """
        try:
            # Extract document ID from URL
            doc_id = self._extract_doc_id(doc_url)
            
            # Initialize Docs service
            service = self._initialize_docs_service()
            
            # Get the document
            document = service.documents().get(documentId=doc_id).execute()
            
            # Extract text content from paragraphs, tables, and TOC
            body_content = document.get('body', {}).get('content', [])
            full_content = self._extract_text_from_elements(body_content)

            if not full_content.strip():
                return "No content found in the document"

            return full_content
            
        except ValueError as e:
            return str(e)
        except Exception as e:
            return f"Error reading Google Doc: {str(e)}"

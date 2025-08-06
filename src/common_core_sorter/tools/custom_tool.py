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
    range: str = Field(description="The range to write to (e.g. 'Sheet1!A1:D10')")
    values: List[List[str]] = Field(description="The values to write to the sheet")


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

    def _run(self, spreadsheet_id: str, range: str, values: List[List[str]]) -> str:
        """
        Write data to Google Sheets
        
        Args:
            spreadsheet_id: The ID of the Google Spreadsheet
            range: The range to write to
            values: The values to write
            
        Returns:
            Status message
        """
        try:
            service = self._initialize_sheets_service()
            sheet = service.spreadsheets()
            
            # Append the data to the sheet
            result = sheet.values().append(
                spreadsheetId=spreadsheet_id,
                range=range,
                valueInputOption='RAW',
                insertDataOption='INSERT_ROWS',
                body={'values': values}
            ).execute()
            
            updates = result.get('updates', {})
            rows_updated = updates.get('updatedRows', 0)
            
            return f"Successfully wrote {rows_updated} rows to Google Sheets"
            
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
            
            # Extract text content
            content = []
            for element in document.get('body', {}).get('content', []):
                if 'paragraph' in element:
                    for text_element in element['paragraph'].get('elements', []):
                        if 'textRun' in text_element:
                            content.append(text_element['textRun'].get('content', ''))
            
            full_content = ''.join(content)
            
            if not full_content:
                return "No content found in the document"
            
            return full_content
            
        except ValueError as e:
            return str(e)
        except Exception as e:
            return f"Error reading Google Doc: {str(e)}"
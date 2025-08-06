#!/usr/bin/env python3
"""
Extract lesson plan URLs from Google Drive folders listed in CSV

This script reads a CSV file containing lesson information, accesses Google Drive
folders from column F, finds lesson plan documents, and outputs a JSON mapping
of lesson IDs (column A) to lesson plan URLs.

Author: Assistant
Date: 2025-01-28
"""

import csv
import json
import re
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
from typing import Optional, Dict, List


def extract_folder_id(url: str) -> Optional[str]:
    """Extract folder ID from Google Drive URL"""
    if not url or not isinstance(url, str):
        return None
    
    # Patterns for different Drive URL formats
    patterns = [
        r'/folders/([a-zA-Z0-9_-]+)',
        r'/d/([a-zA-Z0-9_-]+)',
        r'id=([a-zA-Z0-9_-]+)'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    
    return None


def initialize_drive_service(credentials_path: str):
    """Initialize Google Drive API service"""
    credentials = service_account.Credentials.from_service_account_file(
        credentials_path,
        scopes=[
            'https://www.googleapis.com/auth/drive.readonly',
            'https://www.googleapis.com/auth/drive.metadata.readonly'
        ]
    )
    return build('drive', 'v3', credentials=credentials)


def find_lesson_plan_in_folder(service, folder_id: str) -> Optional[str]:
    """
    Find lesson plan document in a Google Drive folder
    Returns the shareable link URL if found
    """
    try:
        # List all files in the folder
        query = f"'{folder_id}' in parents and trashed = false"
        results = service.files().list(
            q=query,
            fields="files(id, name, mimeType, webViewLink)",
            pageSize=100,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        
        files = results.get('files', [])
        
        # Debug: Print all files found
        if files:
            print(f"  Found {len(files)} files in folder:")
            for file in files:  # Show all files for debugging
                print(f"    - {file.get('name')} ({file.get('mimeType')})")
        else:
            print(f"  Folder appears to be empty or not accessible")
            print(f"  Trying without trashed filter...")
            # Try without trashed filter
            query = f"'{folder_id}' in parents"
            results = service.files().list(
                q=query,
                fields="files(id, name, mimeType, webViewLink)",
                pageSize=100,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True
            ).execute()
            files = results.get('files', [])
            if files:
                print(f"  Found {len(files)} files without filter")
        
        # Look for files following the naming convention: FINALIZED *_LessonPlan
        for file in files:
            file_name = file.get('name', '')
            file_name_lower = file_name.lower()
            
            # Check for FINALIZED lesson plan pattern
            if 'finalized' in file_name_lower and 'lessonplan' in file_name_lower:
                print(f"  ✓ Found FINALIZED lesson plan: {file.get('name')}")
                return file.get('webViewLink')
        
        # Look for files with "lessonplan" in the name (case-insensitive)
        for file in files:
            file_name = file.get('name', '').lower()
            if 'lessonplan' in file_name or 'lesson plan' in file_name or 'lesson_plan' in file_name:
                # Check if it's a Google Doc or other document type
                mime_type = file.get('mimeType', '')
                if 'document' in mime_type or 'word' in mime_type or 'text' in mime_type or 'pdf' in mime_type:
                    print(f"  ✓ Found lesson plan: {file.get('name')}")
                    return file.get('webViewLink')
        
        # If no exact match, look for any doc with "lesson" in the name
        for file in files:
            file_name = file.get('name', '').lower()
            if 'lesson' in file_name:
                mime_type = file.get('mimeType', '')
                if 'document' in mime_type or 'word' in mime_type or 'pdf' in mime_type:
                    print(f"  ✓ Found lesson document: {file.get('name')}")
                    return file.get('webViewLink')
        
        # Try looking for any Google Doc or Word document
        for file in files:
            mime_type = file.get('mimeType', '')
            if 'google-apps.document' in mime_type or 'word' in mime_type:
                print(f"  ✓ Found document (possible lesson plan): {file.get('name')}")
                return file.get('webViewLink')
        
        print(f"  No lesson plan found among {len(files)} files")
        return None
        
    except Exception as e:
        error_msg = str(e)
        print(f"  Error accessing folder {folder_id}: {error_msg}")
        
        # Check for specific permission errors
        if '403' in error_msg or 'forbidden' in error_msg.lower():
            print(f"  ⚠️  Permission denied. Make sure the folder is shared with the service account.")
        elif '404' in error_msg:
            print(f"  ⚠️  Folder not found. The folder may have been deleted or the ID is incorrect.")
        
        return None


def main():
    # Paths
    csv_path = '/Users/ian/Desktop/Code 2/SOCS4all/SOCS4AI/crewai-core-concept-mover/FINALIZED CT Lessons Tracker  - 2024-2025-2.csv'
    credentials_path = '/Users/ian/Desktop/Code 2/SOCS4all/SOCS4AI/crewai-core-concept-mover/socs4all-e896217ba3d5.json'
    output_path = '/Users/ian/Desktop/Code 2/SOCS4all/SOCS4AI/crewai-core-concept-mover/lesson_plan_urls.json'
    
    # Check if files exist
    if not os.path.exists(csv_path):
        print(f"Error: CSV file not found at {csv_path}")
        return
    
    if not os.path.exists(credentials_path):
        print(f"Error: Credentials file not found at {credentials_path}")
        return
    
    print("Initializing Google Drive service...")
    service = initialize_drive_service(credentials_path)
    
    # Print service account email for reference
    with open(credentials_path, 'r') as f:
        creds_data = json.load(f)
        service_account_email = creds_data.get('client_email', 'Not found')
        print(f"\n⚠️  Service Account Email: {service_account_email}")
        print("⚠️  Make sure the Google Drive folders are shared with this email address!")
        print("⚠️  Without proper sharing, the folders will appear empty.\n")
    
    # Read CSV and process
    results = []
    
    print(f"\nReading CSV file: {csv_path}")
    with open(csv_path, 'r', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        headers = next(reader)  # Skip header row
        
        # Find column indices
        col_a_idx = 0  # Lesson ID (column A)
        col_e_idx = 4  # Ready status (column E) - TRUE/FALSE
        col_f_idx = 5  # Link to folder (column F)
        
        print(f"Processing rows...")
        total_rows = 0
        skipped_rows = 0
        
        for row_num, row in enumerate(reader, start=2):
            total_rows += 1
            
            if len(row) > col_f_idx:
                lesson_id = row[col_a_idx].strip()
                ready_status = row[col_e_idx].strip() if len(row) > col_e_idx else ""
                folder_url = row[col_f_idx].strip()
                
                # Skip rows where column E is FALSE
                if ready_status.upper() == "FALSE":
                    print(f"\nRow {row_num}: Lesson ID {lesson_id} - Skipping (Ready status is FALSE)")
                    skipped_rows += 1
                    continue
                
                if lesson_id and folder_url:
                    print(f"\nRow {row_num}: Lesson ID {lesson_id}")
                    print(f"  Folder URL: {folder_url}")
                    
                    # Extract folder ID
                    folder_id = extract_folder_id(folder_url)
                    if folder_id:
                        print(f"  Extracted folder ID: {folder_id}")
                        
                        # Find lesson plan in folder
                        lesson_plan_url = find_lesson_plan_in_folder(service, folder_id)
                        
                        if lesson_plan_url:
                            print(f"  ✓ Found lesson plan URL: {lesson_plan_url}")
                            results.append({
                                lesson_id: lesson_plan_url
                            })
                        else:
                            print(f"  ⚠️  No lesson plan found in folder")
                    else:
                        print(f"  ⚠️  Could not extract folder ID from URL")
        
        print(f"\n📊 Summary: Processed {total_rows} rows, skipped {skipped_rows} rows (Ready = FALSE)")
    
    # Convert results list to a single dictionary
    final_results = {}
    for item in results:
        final_results.update(item)
    
    # Save results to JSON
    print(f"\nSaving results to: {output_path}")
    with open(output_path, 'w', encoding='utf-8') as jsonfile:
        json.dump(final_results, jsonfile, indent=2)
    
    print(f"\nProcessing complete!")
    print(f"Found lesson plans for {len(final_results)} lessons")
    print(f"Results saved to: {output_path}")


if __name__ == "__main__":
    main()
# CrewAI Google Sheets Integration Implementation Guide

## Overview
This guide explains how to implement a CrewAI flow that processes lesson plan documents and writes extracted information to specific columns in Google Sheets based on lesson IDs.

## Architecture Flow
```
lesson_plan_urls.json → CrewAI kickoff_for_each → Agent processes each doc → Writes to Google Sheets
```

---

## Step 1: Create CrewAI Flow in main.py

### 1.1 Load Input Array
```python
import json
from crewai import Crew, Agent, Task
from src.common_core_sorter.tools.custom_tool import GoogleSheetsWriterTool, GoogleDocReaderTool

# Load the lesson plan URLs from JSON
with open('lesson_plan_urls.json', 'r') as f:
    lesson_data = json.load(f)

# Convert to input array format for kickoff_for_each
input_array = [
    {
        "columnA": item["columnA"],
        "docurl": item["docurl"]
    }
    for item in lesson_data
]
```

### 1.2 Define the Crew and kickoff_for_each
```python
# Initialize tools
sheets_writer = GoogleSheetsWriterTool()
doc_reader = GoogleDocReaderTool()

# Define agent
standards_extractor = Agent(
    role='Standards Extractor',
    goal='Extract state standards from lesson plan documents and write them to Google Sheets',
    backstory='You are an expert at reading educational documents and identifying state standards',
    tools=[doc_reader, sheets_writer],
    verbose=True
)

# Define task (template with placeholders)
extract_standards_task = Task(
    description="""
    1. Read the Google Doc at URL: {docurl}
    2. Extract all state standards mentioned in the document
    3. Write the extracted standards to Google Sheets in column P 
       for the row where column A equals {columnA}
    """,
    expected_output='Confirmation that standards were written to the sheet',
    agent=standards_extractor
)

# Create crew
crew = Crew(
    agents=[standards_extractor],
    tasks=[extract_standards_task],
    verbose=True
)

# Execute for each item in the array
results = crew.kickoff_for_each(inputs=input_array)
```

---

## Step 2: Rewrite custom_tool.py GoogleSheetsWriterTool

### 2.1 Enhanced GoogleSheetsWriterTool with Row Lookup
```python
class GoogleSheetsWriterInput(BaseModel):
    """Schema for Google Sheets write operation parameters"""
    spreadsheet_id: str = Field(description="The ID of the Google Spreadsheet")
    column_a_value: str = Field(description="The value in column A to identify the row")
    column_letter: str = Field(description="The column letter to write to (e.g., 'P')")
    value: str = Field(description="The value to write")

class GoogleSheetsWriterTool(BaseTool):
    """
    Tool for writing data to specific cells in Google Sheets based on column A value.
    """
    name: str = "Google Sheets Writer"
    description: str = "Writes data to a specific cell in Google Sheets by matching column A value"
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
            result = sheet.values().get(
                spreadsheetId=spreadsheet_id,
                range='Sheet1!A:A'  # Adjust sheet name if needed
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
            target_range = f'Sheet1!{column_letter}{row_number}'
            
            # Write the value to the target cell
            body = {
                'values': [[value]]
            }
            
            result = sheet.values().update(
                spreadsheetId=spreadsheet_id,
                range=target_range,
                valueInputOption='RAW',
                body=body
            ).execute()
            
            return f"Successfully wrote to cell {target_range} (row where column A = '{column_a_value}')"
            
        except Exception as e:
            return f"Error writing to Google Sheets: {str(e)}"
```

---

## Step 3: Update CrewAI Task Instructions

### 3.1 Detailed Task Definition
```python
extract_standards_task = Task(
    description="""
    You need to extract state standards from a lesson plan and write them to Google Sheets.
    
    Input provided:
    - columnA: {columnA} (this is the lesson ID from column A in the spreadsheet)
    - docurl: {docurl} (this is the Google Doc URL containing the lesson plan)
    
    Your steps:
    1. Use the Google Doc Reader tool to read the content from {docurl}
    2. Look for sections that mention:
       - State standards
       - Common Core standards
       - Learning standards
       - Any standards referenced (e.g., "CCSS.ELA-LITERACY.W.5.1")
    3. Extract all standards mentioned and format them as a comma-separated list
    4. Use the Google Sheets Writer tool with these parameters:
       - spreadsheet_id: "YOUR_SPREADSHEET_ID_HERE"
       - column_a_value: {columnA}
       - column_letter: "P"
       - value: [the comma-separated list of standards you extracted]
    
    Make sure to handle cases where:
    - No standards are found (write "No standards found")
    - Multiple standards exist (separate with commas)
    - Standards are in different formats
    """,
    expected_output='Confirmation message showing the standards were written to column P',
    agent=standards_extractor
)
```

### 3.2 Multiple Column Updates
If you need to write to multiple columns (e.g., standards to P, objectives to Q, etc.):

```python
comprehensive_extraction_task = Task(
    description="""
    Extract multiple pieces of information from the lesson plan and write to different columns:
    
    For lesson ID {columnA} from document {docurl}:
    
    1. Extract STATE STANDARDS → Write to column P
    2. Extract LEARNING OBJECTIVES → Write to column Q  
    3. Extract ASSESSMENT METHODS → Write to column R
    4. Extract COMPUTATIONAL THINKING CONCEPTS → Write to column S
    
    Use the Google Sheets Writer tool multiple times, once for each column:
    - Always use column_a_value: {columnA} to identify the correct row
    - Change column_letter for each piece of data (P, Q, R, S)
    """,
    expected_output='Confirmation of all data written to respective columns',
    agent=standards_extractor
)
```

---

## Complete Example Implementation

### main.py
```python
import json
import os
from crewai import Crew, Agent, Task
from src.common_core_sorter.tools.custom_tool import GoogleSheetsWriterTool, GoogleDocReaderTool

def main():
    # Configuration
    SPREADSHEET_ID = "your_spreadsheet_id_here"  # Replace with actual ID
    
    # Load lesson URLs
    with open('lesson_plan_urls.json', 'r') as f:
        lesson_data = json.load(f)
    
    # Prepare input array
    input_array = [
        {
            "columnA": item["columnA"],
            "docurl": item["docurl"],
            "spreadsheet_id": SPREADSHEET_ID
        }
        for item in lesson_data
    ]
    
    # Initialize tools
    sheets_writer = GoogleSheetsWriterTool()
    doc_reader = GoogleDocReaderTool()
    
    # Create agent
    standards_extractor = Agent(
        role='Educational Standards Specialist',
        goal='Extract and categorize educational standards from lesson plans',
        backstory='''You are an expert in educational curriculum with deep knowledge 
        of state standards, Common Core standards, and computational thinking concepts. 
        You can quickly identify and extract standards from lesson plan documents.''',
        tools=[doc_reader, sheets_writer],
        verbose=True,
        max_iter=3
    )
    
    # Define task
    extraction_task = Task(
        description="""
        Process lesson plan document and update spreadsheet:
        
        1. Read the Google Doc from URL: {docurl}
        2. Extract the following information:
           - State/Common Core Standards
           - Computational Thinking concepts used
           - Grade level
           - Subject area
        
        3. Write extracted data to spreadsheet (ID: {spreadsheet_id}):
           - Standards → Column P
           - CT Concepts → Column S
           - Grade Level → Column I (if not already filled)
           - Subject → Column K (if not already filled)
        
        For each write operation, use column_a_value: {columnA}
        
        If information is not found, write "Not specified" instead of leaving blank.
        """,
        expected_output='Summary of all information extracted and written to spreadsheet',
        agent=standards_extractor
    )
    
    # Create and run crew
    crew = Crew(
        agents=[standards_extractor],
        tasks=[extraction_task],
        verbose=True
    )
    
    # Process each lesson plan
    print(f"Processing {len(input_array)} lesson plans...")
    results = crew.kickoff_for_each(inputs=input_array)
    
    # Save results
    with open('extraction_results.json', 'w') as f:
        json.dump(results, f, indent=2)
    
    print("Processing complete! Results saved to extraction_results.json")

if __name__ == "__main__":
    main()
```

---

## Configuration Requirements

### Environment Variables (.env)
```bash
GOOGLE_SHEETS_CREDENTIALS=socs4all-e896217ba3d5.json
SPREADSHEET_ID=your_actual_spreadsheet_id_here
```

### Required Files
1. `lesson_plan_urls.json` - Array of lesson IDs and document URLs
2. `socs4all-e896217ba3d5.json` - Google service account credentials
3. `.env` - Environment variables

### Google Sheets Setup
1. Share the spreadsheet with service account email: `socs4all@socs4all.iam.gserviceaccount.com`
2. Ensure Column A contains the lesson IDs (200, 201, 202, etc.)
3. Prepare columns P, Q, R, S for data insertion

---

## Testing Instructions

### Test Single Item First
```python
# Test with just one item before running full batch
test_input = [{
    "columnA": "200",
    "docurl": "https://docs.google.com/document/d/...",
    "spreadsheet_id": SPREADSHEET_ID
}]

results = crew.kickoff_for_each(inputs=test_input)
print(results)
```

### Error Handling
The tool includes error handling for:
- Missing column A values
- API access issues  
- Document read failures
- Sheet write permissions

---

## Notes for Developers

1. **Rate Limiting**: Google APIs have quotas. Consider adding delays between operations if processing many documents.

2. **Batch Operations**: For better performance, consider modifying the tool to support batch updates instead of individual cell writes.

3. **Validation**: Add validation to ensure extracted data meets expected formats before writing to sheets.

4. **Logging**: Implement comprehensive logging to track which documents have been processed.

5. **Resume Capability**: Save progress periodically to allow resuming if the process is interrupted.

6. **Column Mapping**: Consider making column assignments configurable rather than hard-coded.
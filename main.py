import os
import sys
import json
from dotenv import load_dotenv
from crewai import Crew, Agent, Task

# Ensure we can import from src/
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
SRC_PATH = os.path.join(PROJECT_ROOT, 'src')
if SRC_PATH not in sys.path:
    sys.path.insert(0, SRC_PATH)

from common_core_sorter.tools.custom_tool import GoogleSheetsWriterTool, GoogleDocReaderTool


def main():
    load_dotenv()

    spreadsheet_id = os.getenv('SPREADSHEET_ID')
    if not spreadsheet_id:
        raise ValueError("SPREADSHEET_ID is not set in the environment (.env)")

    with open('lesson_plan_urls.json', 'r', encoding='utf-8') as f:
        lesson_data = json.load(f)

    input_array = [
        {
            'columnA': item['columnA'],
            'docurl': item['docurl'],
            'spreadsheet_id': spreadsheet_id,
        }
        for item in lesson_data
    ]

    sheets_writer = GoogleSheetsWriterTool()
    doc_reader = GoogleDocReaderTool()

    standards_extractor = Agent(
        role='Educational Standards Specialist',
        goal='Extract and categorize educational standards from lesson plans',
        backstory=(
            'You are an expert in educational curriculum with deep knowledge of '
            'state standards, Common Core standards, and computational thinking concepts.'
        ),
        tools=[doc_reader, sheets_writer],
        verbose=True,
        max_iter=3,
    )

    extraction_task = Task(
        description=(
            """
            Process the lesson plan document and update the spreadsheet.

            1. Use the Google Doc Reader tool to read the content from {docurl}.
            2. Identify all standards mentioned in the document, including:
               - State standards
               - Common Core standards (e.g., "CCSS.ELA-LITERACY.W.5.1")
               - Learning standards
            3. Produce a comma-separated list of all standards found (no extra commentary).
            4. Use the Google Sheets Writer tool with:
               - spreadsheet_id: {spreadsheet_id}
               - column_a_value: {columnA}
               - column_letter: "P"
               - value: [the comma-separated standards list]

            If no standards are found, write "No standards found".
            """
        ),
        expected_output='Confirmation that standards were written to column P for the specified lesson ID',
        agent=standards_extractor,
    )

    crew = Crew(agents=[standards_extractor], tasks=[extraction_task], verbose=True)

    print(f"Processing {len(input_array)} lesson plans...")
    results = crew.kickoff_for_each(inputs=input_array)

    with open('extraction_results.json', 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2)

    print('Processing complete! Results saved to extraction_results.json')


if __name__ == '__main__':
    main() 
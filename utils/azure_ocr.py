import os
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv

load_dotenv()

class AzureOCR:
    def __init__(self):
        endpoint = os.getenv("DOCUMENT_INTELLIGENCE_ENDPOINT")
        key = os.getenv("DOCUMENT_INTELLIGENCE_KEY")
        
        if not endpoint or not key:
            raise ValueError("Missing DOCUMENT_INTELLIGENCE_ENDPOINT or DOCUMENT_INTELLIGENCE_KEY in .env file")
            
        self.model_id = os.getenv("DI_MODEL_ID", "prebuilt-layout")
        self.client = DocumentAnalysisClient(endpoint, AzureKeyCredential(key))

    def analyze_document(self, file_path):
        """Analyzes a document using Azure Document Intelligence Layout model."""
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
            
        with open(file_path, "rb") as f:
            poller = self.client.begin_analyze_document(self.model_id, document=f)
            result = poller.result()
        
        return result

    def extract_data(self, result):
        """Extracts structured text, tables, and checkboxes from the result."""
        extracted_data = {
            "content": result.content if hasattr(result, 'content') else "",
            "tables": [],
            "selection_marks": [],
            "pages": []
        }

        # Extract Tables
        if hasattr(result, 'tables') and result.tables:
            for table in result.tables:
                table_data = []
                rows = [[] for _ in range(table.row_count)]
                for cell in table.cells:
                    rows[cell.row_index].append({
                        "content": cell.content,
                        "row_index": cell.row_index,
                        "column_index": cell.column_index
                    })
                
                # Sort cells by column index for each row
                for row in rows:
                    row.sort(key=lambda x: x["column_index"])
                    table_data.append([cell["content"] for cell in row])
                
                extracted_data["tables"].append(table_data)

        # Extract Pages, Lines, and Selection Marks (Checkboxes)
        if hasattr(result, 'pages'):
            for page in result.pages:
                page_marks = []
                if hasattr(page, 'selection_marks') and page.selection_marks:
                    for mark in page.selection_marks:
                        page_marks.append({
                            "state": mark.state, # 'selected' or 'unselected'
                            "polygon": getattr(mark, 'polygon', None),
                            "confidence": getattr(mark, 'confidence', 0),
                            "page_number": page.page_number
                        })
                
                extracted_data["selection_marks"].append(page_marks)
                extracted_data["pages"].append({
                    "page_number": page.page_number,
                    "lines": [line.content for line in page.lines] if hasattr(page, 'lines') else []
                })

        return extracted_data

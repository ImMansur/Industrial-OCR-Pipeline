import os
import sys
from utils.azure_ocr import AzureOCR
from utils.azure_openai import OpenAIStructurer
from utils.excel_exporter import ExcelExporter
from dotenv import load_dotenv

load_dotenv()

def process_document(file_path):
    print(f"--- Processing: {os.path.basename(file_path)} ---")
    
    # 1. OCR Extraction
    ocr = AzureOCR()
    print("Extracting data via Azure Document Intelligence...")
    raw_result = ocr.analyze_document(file_path)
    extracted_data = ocr.extract_data(raw_result)
    
    # 2. LLM Structuring
    structurer = OpenAIStructurer()
    print("Structuring text via Azure OpenAI...")
    structured_info = structurer.structure_text(extracted_data["content"])
    
    # Add metadata
    structured_info["Filename"] = os.path.basename(file_path)
    extracted_data["structured_info"] = structured_info
    
    # 3. Individual Excel Export (Optional/Backup)
    output_filename = os.path.splitext(os.path.basename(file_path))[0] + "_extracted.xlsx"
    output_path = os.path.join("data", "output", output_filename)
    
    exporter = ExcelExporter()
    exporter.export(extracted_data, output_path)
    print(f"--- Finished: {output_filename} ---\n")
    
    return {
        "structured_info": structured_info,
        "tables": extracted_data.get("tables", []),
        "filename": os.path.basename(file_path)
    }

def main():
    input_dir = os.path.join("data", "input")
    files = [f for f in os.listdir(input_dir) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg'))]
    
    if not files:
        print(f"No documents found in {input_dir}. Please add some files to process.")
        return

    all_results = []
    for file in files:
        file_path = os.path.join(input_dir, file)
        try:
            result = process_document(file_path)
            all_results.append(result)
        except Exception as e:
            print(f"Error processing {file}: {e}")

    # Generate Consolidated Report with 3 Sheets
    if all_results:
        import pandas as pd
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        
        consolidated_path = os.path.join("data", "output", "consolidated_ocr_report.xlsx")
        
        with pd.ExcelWriter(consolidated_path, engine='openpyxl') as writer:
            # Sheet 1: Certificate Details (Header Information from all certificates)
            cert_details = [result["structured_info"] for result in all_results]
            df_certificates = pd.DataFrame(cert_details)
            cols = ['Filename'] + [c for c in df_certificates.columns if c != 'Filename']
            df_certificates = df_certificates[cols]
            df_certificates.to_excel(writer, sheet_name="Certificate Details", index=False)
            
            # Sheet 2: Equipment Line Items (All items from all certificates)
            all_equipment = []
            for result in all_results:
                filename = result["filename"]
                cert_id = result["structured_info"].get("Equipment Serial Number", 
                         result["structured_info"].get("Certification Details", filename))
                
                for table in result.get("tables", []):
                    df_table = pd.DataFrame(table)
                    if not df_table.empty:
                        # Add source certificate reference
                        df_table.insert(0, "Certificate", cert_id)
                        df_table.insert(1, "Source File", filename)
                        all_equipment.append(df_table)
            
            if all_equipment:
                df_equipment = pd.concat(all_equipment, ignore_index=True)
                df_equipment.to_excel(writer, sheet_name="Equipment Line Items", index=False)
            
            # Sheet 3: Consolidated Register (All parts with certificate references)
            if all_equipment:
                df_consolidated = pd.concat([df for df in all_equipment], ignore_index=True)
                # Sort by part number if available
                if len(df_consolidated.columns) > 3:
                    try:
                        df_consolidated = df_consolidated.sort_values(by=df_consolidated.columns[2])
                    except:
                        pass
                df_consolidated.to_excel(writer, sheet_name="Consolidated Register", index=False)
            
            # Format all sheets
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                # Apply formatting
                header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # Format headers
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = border
                
                # Format data cells
                for row in worksheet.iter_rows(min_row=2):
                    for cell in row:
                        cell.border = border
                        cell.alignment = Alignment(vertical="top", wrap_text=True)
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                worksheet.freeze_panes = "A2"
        
        print(f"\nCONSOLIDATED REPORT GENERATED: {consolidated_path}")
        print(f"  ✓ Sheet 1: Certificate Details ({len(cert_details)} certificates)")
        print(f"  ✓ Sheet 2: Equipment Line Items")
        print(f"  ✓ Sheet 3: Consolidated Register")

if __name__ == "__main__":
    main()

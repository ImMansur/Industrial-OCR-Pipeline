import os
from utils.azure_ocr import AzureOCR

ocr = AzureOCR()
file_path = os.path.join('data', 'input', 'Pages from DB BAW379-double.pdf')
result = ocr.analyze_document(file_path)
extracted = ocr.extract_data(result)
print('tables count', len(extracted['tables']))
for ti, table in enumerate(extracted['tables']):
    print('--- TABLE', ti, 'rows', len(table))
    for ri, row in enumerate(table[:12]):
        print(ri, row)
    print()
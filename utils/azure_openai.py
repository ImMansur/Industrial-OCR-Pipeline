import os
import json
from openai import AzureOpenAI
from dotenv import load_dotenv

load_dotenv()

class OpenAIStructurer:
    def __init__(self):
        endpoint = os.getenv("AZURE_OPENAI_ENDPOINT")
        api_key = os.getenv("AZURE_OPENAI_KEY")
        api_version = os.getenv("AZURE_OPENAI_API_VERSION")
        self.deployment_name = os.getenv("AZURE_OPENAI_DEPLOYMENT")

        if not all([endpoint, api_key, api_version, self.deployment_name]):
            raise ValueError("Missing Azure OpenAI credentials (endpoint, key, version, or deployment) in .env file")

        self.client = AzureOpenAI(
            azure_endpoint=endpoint,
            api_key=api_key,
            api_version=api_version
        )

    def structure_text(self, text_content):
        """Uses Azure OpenAI to extract specific structured fields from OCR text."""
        prompt = f"""
        Extract the following fields from the Certificate of Conformity OCR text provided below. 
        If a field is not found, return null.
        Return the output strictly in JSON format.

        Fields to extract:
        1. Document Type (e.g., "Certificate of Conformity", "Consolidated Parts Register")
        2. Issuer (Company name issuing the certificate)
        3. Address (Full address of issuer)
        4. Phone (Phone number)
        5. Fax (Fax number)
        6. Certificate Date
        7. Customer Name
        8. Purchase Order (PO number)
        9. Sales Order (SO number)
        10. Serialization (Main certificate serialization/batch number, e.g., BAW106, BS126)
        11. Applicable Specs (e.g., API 6A, API 16A)
        12. Authorized Signatory (Name of person who signed)
        13. Signatory Title (Title/position of signatory)
        14. Total Items (Total number of equipment items listed)

        OCR Text:
        {text_content}
        """

        response = self.client.chat.completions.create(
            model=self.deployment_name,
            messages=[
                {"role": "system", "content": "You are a helpful assistant that extracts structured data from Certificate of Conformity documents into JSON format. Focus on accuracy and completeness."},
                {"role": "user", "content": prompt}
            ],
            response_format={"type": "json_object"}
        )

        return json.loads(response.choices[0].message.content)

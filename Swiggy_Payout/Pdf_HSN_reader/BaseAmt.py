import pdfplumber
import pandas as pd
import re
import os

results = []

for file in os.listdir("."):

    if not file.lower().endswith(".pdf"):
        continue

    try:
        text = ""

        # Read PDF
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        # -----------------------------
        # Extract Invoice Number
        # -----------------------------
        invoice_no = ""

        patterns = [
            r"Invoice\s*Number\s*:?\s*([A-Z0-9\-]+)",
            r"Invoice\s*No\.?\s*:?\s*([A-Z0-9\-]+)",
            r"DocDtls_No\s*:?\s*([A-Z0-9\-]+)"
        ]

        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                invoice_no = m.group(1).strip()
                break

        # Fallback: get invoice number from filename
        if not invoice_no:
            m = re.search(r'(\d{6}FS\d+)', file)
            if m:
                invoice_no = m.group(1)

        if not invoice_no:
            print(f"Invoice number not found: {file}")

        # -----------------------------
        # HSN Calculation
        # -----------------------------
        ending_11_total = 0
        other_total = 0

        for line in text.split("\n"):

            match = re.search(
                r'(\d{6})\s+\w+\s+\d+\s+([\d,]+\.\d+|[\d,]+)',
                line
            )

            if match:

                hsn = match.group(1)

                nums = re.findall(
                    r'[\d,]+\.\d+|[\d,]+',
                    line
                )

                if len(nums) >= 4:

                    try:
                        base_amount = float(
                            nums[3].replace(",", "")
                        )

                        if hsn.endswith("11"):
                            ending_11_total += base_amount
                        else:
                            other_total += base_amount

                    except:
                        pass

        results.append({
            "Invoice No": invoice_no,
            "PDF File": file,
            "HSN Ending 11": round(ending_11_total, 2),
            "HSN Not Ending 11": round(other_total, 2),
            "Subtotal": round(
                ending_11_total + other_total,
                2
            )
        })

        print(f"Processed: {file}")

    except Exception as e:
        print(f"Error in {file}: {e}")

# Save Excel
df = pd.DataFrame(results)
df.to_excel("hsn_summary.xlsx", index=False)

print("\nDone!")
print("Output saved to hsn_summary.xlsx")
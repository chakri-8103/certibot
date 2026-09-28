import generator
import time

print("Starting batch PDF generation across all sheets in Excel...")
start_time = time.time()

res = generator.process_excel_and_generate_all(
    template_path="250371509001.pdf",
    excel_path="I SEM DIGIVAL PURPOSE-25 AB F+RV.xlsx",
    signature_folder="signatures",
    output_base_dir="generated_pdfs",
    selected_sheets=None,  # All sheets!
    date_mode="fixed",
    fixed_date="06.04.2026"
)

elapsed = time.time() - start_time
print(f"\n==========================================")
print(f"BATCH GENERATION COMPLETE in {elapsed:.2f} seconds")
print(f"==========================================")
print(f"Total Students Processed: {res['total']}")
print(f"PDFs Successfully Generated: {res['generated']}")
print(f"Failed: {res['failed']}")
print(f"Zip Archive Created At: {res['zip_path']}")

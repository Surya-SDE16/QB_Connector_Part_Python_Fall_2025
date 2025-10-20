from excel_processor import process_payment_terms
from item_list_processor import process_items

def main():
    excel_file_path = "path/to/your_excel_file.xlsx"  # update your path
    
    print("=== Starting Payment Terms Comparison ===")
    try:
        payment_terms_comparison = process_payment_terms(excel_file_path)
        print("Payment terms comparison completed.\n")
    except Exception as e:
        print(f"Payment terms comparison failed: {e}")

    print("=== Starting Item List Comparison ===")
    try:
        item_list_comparison = process_items(excel_file_path)
        print("Item list comparison completed.\n")
    except Exception as e:
        print(f"Item list comparison failed: {e}")

if __name__ == "__main__":
    main()

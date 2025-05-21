import argparse
from process import process_data_type

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch and update Google Sheets data.")
    parser.add_argument("type", choices=["issues", "mr", "ntc"], help="Type of data to process")
    args = parser.parse_args()

    process_data_type(args.type)


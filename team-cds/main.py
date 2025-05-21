import argparse
from process import process_data_type

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch and update Google Sheets data.")
    parser.add_argument("type", choices=["issues", "mr", "ntc"], help="Type of data to process")
    parser.add_argument(
        "--utils_range",
        default="UTILS!B2:B1000",
        help="Google Sheets range for UTILS sheet (default: 'UTILS!B2:B1000')"
    )
    args = parser.parse_args()

    # Pass the fixed range string to process_data_type if it accepts it
    process_data_type(args.type, utils_range=args.utils_range)

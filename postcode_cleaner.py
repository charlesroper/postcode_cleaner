import pandas as pd
import argparse
import postcodes_uk
from postcodes_uk import Postcode


def clean_postcode(postcode):
    if pd.isna(postcode):
        return None  # Handle NaN values
    postcode = postcode.strip().upper()  # Standardize input
    if postcodes_uk.validate(postcode):
        print(postcode)
        return Postcode.from_string(postcode)
    else:
        return None  # or return "INVALID" to mark invalid postcodes


def main(input_file, postcode_column):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)

    # Ensure the specified postcode column exists in the DataFrame
    if postcode_column in df.columns:
        # Apply the cleaning function to create a new column named POSTCODE_CLEAN
        df["POSTCODE_CLEAN"] = df[postcode_column].apply(clean_postcode)
    else:
        print(f"Column '{postcode_column}' not found in the Excel file.")
        return

    # Save the updated DataFrame to a new Excel file
    output_file = "output.xlsx"
    df.to_excel(output_file, index=False)
    print(
        f"Data cleansing complete. The updated file has been saved as '{output_file}'."
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Clean and validate UK postcodes in an Excel file."
    )
    parser.add_argument(
        "input_file", type=str, help="The input Excel file to be processed."
    )
    parser.add_argument(
        "postcode_column",
        type=str,
        help="The name of the postcode column to be cleaned.",
    )
    args = parser.parse_args()

    main(args.input_file, args.postcode_column)

import argparse
import pandas as pd
import re
import tabulate
import openpyxl


DEFAULT_IGNORE_SHEETS = {"Key", "Gardening", "Holotactics"}


def main(file_path, output_path, ignore_sheets):
    xls = pd.ExcelFile(file_path)

    sections = []

    for sheet_name in xls.sheet_names:
        if sheet_name in ignore_sheets:
            continue

        df = xls.parse(sheet_name).iloc[:, :3]
        df.columns = ["LOCATION", "ROUTE", "NOTES"]
        df = df.fillna("")

        current_location = None
        location_rows = []

        for _, row in df.iterrows():
            loc = row["LOCATION"].strip()
            route = row["ROUTE"]
            notes = row["NOTES"]

            if loc:
                # skip ETA pseudo-locations
                if loc.upper().startswith("ETA"):
                    continue
                if current_location:
                    sections.append((current_location, location_rows))
                current_location = loc
                location_rows = []

            if current_location and route:
                location_rows.append((route, notes))

        if current_location and location_rows:
            sections.append((current_location, location_rows))

    markdown_blocks = ["\n"]

    for location, rows in sections:
        markdown_blocks.append(f"## {location}")
        table_df = pd.DataFrame(rows, columns=["ROUTE", "NOTES"])
        markdown_blocks.append(table_df.to_markdown(index=False))

    final_markdown = "\n\n".join(markdown_blocks)

    final_markdown = re.sub(r'(\n## .+?)\n\n', r'\1\n', final_markdown)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(final_markdown)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert Excel routes file to Markdown"
    )

    parser.add_argument(
        "--input",
        required=True,
        help="Path to input Excel file"
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Path to output Markdown file"
    )
    parser.add_argument(
        "--ignore-sheets",
        help="Optional comma-separated list of sheet names to ignore (Default is: Key, Gardening, Holotactics)"
    )

    args = parser.parse_args()

    if args.ignore_sheets:
        ignore_sheets = {
            name.strip() for name in args.ignore_sheets.split(",") if name.strip()
        }
    else:
        ignore_sheets = DEFAULT_IGNORE_SHEETS

    main(args.input, args.output, ignore_sheets)

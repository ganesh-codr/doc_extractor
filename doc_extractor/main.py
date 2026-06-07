import argparse

from .extractor import build_schedule_rows, export_excel, extract_text


def main():
    parser = argparse.ArgumentParser(description="Extract drawing information into Excel")
    parser.add_argument("--input", required=True, help="Path to a PDF or text file")
    parser.add_argument("--prompt", required=True, help="User prompt describing the extraction target")
    parser.add_argument("--output", default="output.xlsx", help="Output Excel path")
    args = parser.parse_args()

    source_type = "text" if args.input.endswith(".txt") else "pdf"
    text = extract_text(args.input, source_type=source_type)
    rows = build_schedule_rows(text, args.prompt)
    output_path = export_excel(rows, args.output)

    print(f"Extraction completed. Results written to {output_path}")
    print(f"Rows generated: {len(rows)}")


if __name__ == "__main__":
    main()

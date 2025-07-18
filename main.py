import argparse
from prepTM import extract_som_entries


if __name__ == "__main__":

    parser = argparse.ArgumentParser(
        description="Extract and format suspicious object entries from an EML for Trend Micro SOM"
    )
    parser.add_argument('eml_file', help="Path to the .eml file to process")
    parser.add_argument(
        '--desc', default=None,
        help="Description prefix (defaults to cert+<today>)"
    )
    parser.add_argument(
        '--output', default='TM_filled.csv',
        help="Output CSV path"
    )
    args = parser.parse_args()

    df = extract_som_entries(args.eml_file, args.desc)
    df.to_csv(args.output, index=False)
    print(f"Saved {len(df)} entries to {args.output}")
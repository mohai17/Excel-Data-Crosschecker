import pandas as pd
from pathlib import Path
import sys

if len(sys.argv) != 4:
    print("Usage:")
    print("python compare_excel.py file1.xlsx file2.xlsx output.xlsx")
    sys.exit(1)

FILE1 = sys.argv[1]
FILE2 = sys.argv[2]
OUTPUT = sys.argv[3]


def normalize_dataframe(df):
    df = df.fillna("")
    df.columns = [str(col).strip() for col in df.columns]

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def compare_sheets(writer, sheet_name, df1, df2):
    results = []

    df1 = normalize_dataframe(df1)
    df2 = normalize_dataframe(df2)

    # Row count check
    if len(df1) != len(df2):
        results.append({
            "Issue": "Row Count Mismatch",
            "File1 Rows": len(df1),
            "File2 Rows": len(df2)
        })

    # Column check
    cols1 = set(df1.columns)
    cols2 = set(df2.columns)

    missing_in_file2 = cols1 - cols2
    missing_in_file1 = cols2 - cols1

    for col in missing_in_file2:
        results.append({
            "Issue": "Column Missing in File2",
            "Column": col
        })

    for col in missing_in_file1:
        results.append({
            "Issue": "Column Missing in File1",
            "Column": col
        })

    common_cols = list(cols1.intersection(cols2))

    df1 = df1[common_cols]
    df2 = df2[common_cols]

    max_rows = max(len(df1), len(df2))

    df1 = df1.reindex(range(max_rows)).fillna("")
    df2 = df2.reindex(range(max_rows)).fillna("")

    # Cell comparison
    for row_idx in range(max_rows):

        for col in common_cols:

            val1 = str(df1.at[row_idx, col])
            val2 = str(df2.at[row_idx, col])

            if val1 != val2:
                results.append({
                    "Issue": "Cell Difference",
                    "Row": row_idx + 2,
                    "Column": col,
                    "File1 Value": val1,
                    "File2 Value": val2
                })

    if not results:
        results.append({
            "Issue": "No Differences Found"
        })

    result_df = pd.DataFrame(results)

    safe_sheet_name = sheet_name[:31]

    result_df.to_excel(
        writer,
        sheet_name=safe_sheet_name,
        index=False
    )

    return {
        "Sheet": sheet_name,
        "Differences": len(result_df[result_df["Issue"] != "No Differences Found"])
    }


def main():

    excel1 = pd.ExcelFile(FILE1)
    excel2 = pd.ExcelFile(FILE2)

    summary = []

    with pd.ExcelWriter(
        OUTPUT,
        engine="openpyxl"
    ) as writer:

        sheets1 = set(excel1.sheet_names)
        sheets2 = set(excel2.sheet_names)

        all_sheets = sorted(sheets1.union(sheets2))

        for sheet in all_sheets:

            if sheet not in sheets1:
                pd.DataFrame([
                    {
                        "Issue": f"Sheet missing in File1: {sheet}"
                    }
                ]).to_excel(
                    writer,
                    sheet_name=sheet[:31],
                    index=False
                )

                summary.append({
                    "Sheet": sheet,
                    "Differences": "Missing in File1"
                })

                continue

            if sheet not in sheets2:
                pd.DataFrame([
                    {
                        "Issue": f"Sheet missing in File2: {sheet}"
                    }
                ]).to_excel(
                    writer,
                    sheet_name=sheet[:31],
                    index=False
                )

                summary.append({
                    "Sheet": sheet,
                    "Differences": "Missing in File2"
                })

                continue

            df1 = pd.read_excel(
                FILE1,
                sheet_name=sheet,
                dtype=str
            )

            df2 = pd.read_excel(
                FILE2,
                sheet_name=sheet,
                dtype=str
            )

            result = compare_sheets(
                writer,
                sheet,
                df1,
                df2
            )

            summary.append(result)

        pd.DataFrame(summary).to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

    print("=" * 50)
    print("Comparison Completed")
    print(f"Report Generated: {OUTPUT}")
    print("=" * 50)


if __name__ == "__main__":
    main()
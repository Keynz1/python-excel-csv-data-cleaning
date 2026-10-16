from pathlib import Path

from merge_validate import deduplicate, normalize_row, read_rows, validate_row, write_clean_csv, write_clean_xlsx, write_issues_report

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "input"
OUTPUT = ROOT / "output"


def main() -> None:
    paths = list(INPUT.glob("*.csv")) + list(INPUT.glob("*.xlsx"))
    raw_rows = read_rows(paths)
    normalized = [normalize_row(row, str(row.get("__source", "")), int(row.get("__row_number", 0) or 0)) for row in raw_rows]
    clean_rows = deduplicate(normalized)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    write_clean_csv(clean_rows, OUTPUT / "clean_merged.csv")
    write_clean_xlsx(clean_rows, OUTPUT / "clean_merged.xlsx")
    write_issues_report(clean_rows, OUTPUT / "issues_report.csv")
    issue_count = sum(len(validate_row(row)) for row in clean_rows)
    duplicate_count = len(normalized) - len(clean_rows)
    report = "\n".join([
        "Merge / Deduplication / Validation Demo Report",
        f"Input rows: {len(normalized)}",
        f"Unique output rows: {len(clean_rows)}",
        f"Duplicates removed: {duplicate_count}",
        f"Validation issues: {issue_count}",
        "Duplicate rule: normalized ID; source filename then row number wins conflicts.",
        "",
    ])
    (OUTPUT / "processing_report.txt").write_text(report, encoding="utf-8")
    print(report)


if __name__ == "__main__":
    main()

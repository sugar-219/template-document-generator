"""Export full check results as an Excel-readable CSV without partial writes."""
import csv
import tempfile
from pathlib import Path

COLUMNS = ("数据文件", "工作表", "行号", "字段", "数据列", "级别", "说明")


def write_check_report(destination, source, sheet, issues):
    destination = Path(destination)
    source = str(Path(source).resolve())
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8-sig", newline="", dir=destination.parent,
                                         prefix=".check-report-", suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            writer = csv.DictWriter(stream, fieldnames=COLUMNS)
            writer.writeheader()
            for issue in issues:
                writer.writerow({"数据文件": source, "工作表": sheet,
                                 **{column: issue[column] for column in COLUMNS[2:]}})
        temporary.replace(destination)
        return destination
    finally:
        if temporary:
            temporary.unlink(missing_ok=True)

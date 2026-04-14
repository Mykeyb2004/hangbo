from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import pandas as pd

from survey_customer_category_rules import CUSTOMER_CATEGORY_RULES

DEFAULT_SHEET_NAME = "问卷数据"
DEFAULT_LOG_DIR = "logs"


@dataclass(frozen=True)
class MissingSourceNotice:
    source_file_name: str
    expected_path: Path


@dataclass(frozen=True)
class SourceIssue:
    source_file_name: str
    workbook_path: Path
    message: str


@dataclass(frozen=True)
class UnmappedRecord:
    source_file_name: str
    workbook_path: Path
    sheet_name: str
    excel_row_number: int
    data_column: str
    data_value: str
    auxiliary_column: str | None
    auxiliary_value: str | None
    reason: str


@dataclass(frozen=True)
class SourceAudit:
    source_file_name: str
    workbook_path: Path
    sheet_name: str
    total_rows: int
    data_column: str
    auxiliary_column: str | None
    unmapped_records: tuple[UnmappedRecord, ...]


@dataclass(frozen=True)
class DirectoryAuditReport:
    input_dir: Path
    sheet_name: str
    generated_at: datetime
    source_audits: tuple[SourceAudit, ...]
    missing_sources: tuple[MissingSourceNotice, ...]
    source_issues: tuple[SourceIssue, ...]

    @property
    def expected_source_file_count(self) -> int:
        return len(expected_source_file_names())

    @property
    def checked_source_file_count(self) -> int:
        return len(self.source_audits)

    @property
    def files_with_unmapped_records(self) -> int:
        return sum(1 for audit in self.source_audits if audit.unmapped_records)

    @property
    def total_unmapped_records(self) -> int:
        return sum(len(audit.unmapped_records) for audit in self.source_audits)


def expected_source_file_names() -> tuple[str, ...]:
    return tuple(dict.fromkeys(rule.source_file_name for rule in CUSTOMER_CATEGORY_RULES))


def rules_for_source_file(source_file_name: str):
    return tuple(
        rule for rule in CUSTOMER_CATEGORY_RULES if rule.source_file_name == source_file_name
    )


def collect_mapped_pairs(rules) -> set[tuple[str, str]]:
    mapped_pairs: set[tuple[str, str]] = set()
    for rule in rules:
        if not rule.auxiliary_values:
            continue
        for auxiliary_value in rule.auxiliary_values:
            for data_value in rule.data_values:
                mapped_pairs.add((auxiliary_value, data_value))
    return mapped_pairs


def collect_mapped_values(rules) -> set[str]:
    return {
        data_value
        for rule in rules
        for data_value in rule.data_values
        if data_value
    }


def excel_column_to_index(column_name: str) -> int:
    normalized = column_name.strip().upper()
    if not normalized:
        raise ValueError("Excel 列号不能为空。")

    value = 0
    for character in normalized:
        if not ("A" <= character <= "Z"):
            raise ValueError(f"非法 Excel 列号: {column_name}")
        value = value * 26 + (ord(character) - ord("A") + 1)
    return value - 1


def load_text_column(
    df: pd.DataFrame,
    column_name: str,
    *,
    column_label: str,
) -> pd.Series:
    column_index = excel_column_to_index(column_name)
    if column_index >= len(df.columns):
        raise ValueError(
            f"来源数据缺少{column_label}列 {column_name}，当前仅有 {len(df.columns)} 列。"
        )

    return (
        df.iloc[:, column_index]
        .astype("string")
        .fillna("")
        .str.strip()
    )


def audit_source_file(
    *,
    source_file_name: str,
    workbook_path: Path,
    sheet_name: str,
    df: pd.DataFrame,
) -> SourceAudit:
    rules = rules_for_source_file(source_file_name)
    if not rules:
        raise ValueError(f"未找到来源文件 {source_file_name} 对应的客户类别规则。")

    data_columns = {rule.data_column for rule in rules if rule.data_column}
    if len(data_columns) != 1:
        raise ValueError(f"{source_file_name} 的数据标签列定义不唯一，无法核查。")
    data_column = next(iter(data_columns))
    data_series = load_text_column(df, data_column, column_label="数据标签")

    auxiliary_columns = {rule.auxiliary_column for rule in rules if rule.auxiliary_column}
    auxiliary_column = None
    auxiliary_series: pd.Series | None = None
    mapped_pairs: set[tuple[str, str]] = set()
    mapped_values: set[str] = set()

    if auxiliary_columns:
        if len(auxiliary_columns) != 1:
            raise ValueError(f"{source_file_name} 的辅助标签列定义不唯一，无法核查。")
        auxiliary_column = next(iter(auxiliary_columns))
        auxiliary_series = load_text_column(df, auxiliary_column, column_label="辅助标签")
        mapped_pairs = collect_mapped_pairs(rules)
    else:
        mapped_values = collect_mapped_values(rules)

    unmapped_records: list[UnmappedRecord] = []
    for row_index, data_value in enumerate(data_series.tolist(), start=2):
        if not data_value:
            continue

        auxiliary_value = None
        reason: str | None = None
        if auxiliary_series is not None and auxiliary_column is not None:
            auxiliary_value = str(auxiliary_series.iloc[row_index - 2]).strip()
            if not auxiliary_value:
                reason = "辅助标签为空，未映射问题：该记录无法匹配客户类别规则。"
            elif (auxiliary_value, data_value) not in mapped_pairs:
                reason = "辅助标签 + 数据标签组合未映射问题：规则中未包含该记录。"
        elif data_value not in mapped_values:
            reason = "数据标签未映射问题：规则中未包含该记录。"

        if reason is None:
            continue

        unmapped_records.append(
            UnmappedRecord(
                source_file_name=source_file_name,
                workbook_path=workbook_path,
                sheet_name=sheet_name,
                excel_row_number=row_index,
                data_column=data_column,
                data_value=data_value,
                auxiliary_column=auxiliary_column,
                auxiliary_value=auxiliary_value,
                reason=reason,
            )
        )

    return SourceAudit(
        source_file_name=source_file_name,
        workbook_path=workbook_path,
        sheet_name=sheet_name,
        total_rows=len(df),
        data_column=data_column,
        auxiliary_column=auxiliary_column,
        unmapped_records=tuple(unmapped_records),
    )


def read_source_dataframe(workbook_path: Path, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(workbook_path, sheet_name=sheet_name)


def run_directory_audit(
    input_dir: Path,
    *,
    sheet_name: str = DEFAULT_SHEET_NAME,
) -> DirectoryAuditReport:
    resolved_input_dir = input_dir.resolve()
    source_audits: list[SourceAudit] = []
    missing_sources: list[MissingSourceNotice] = []
    source_issues: list[SourceIssue] = []

    for source_file_name in expected_source_file_names():
        workbook_path = resolved_input_dir / source_file_name
        if not workbook_path.exists() or not workbook_path.is_file():
            missing_sources.append(
                MissingSourceNotice(
                    source_file_name=source_file_name,
                    expected_path=workbook_path,
                )
            )
            continue

        try:
            df = read_source_dataframe(workbook_path, sheet_name)
            source_audits.append(
                audit_source_file(
                    source_file_name=source_file_name,
                    workbook_path=workbook_path,
                    sheet_name=sheet_name,
                    df=df,
                )
            )
        except Exception as exc:  # pragma: no cover - 防御性兜底
            source_issues.append(
                SourceIssue(
                    source_file_name=source_file_name,
                    workbook_path=workbook_path,
                    message=str(exc),
                )
            )

    return DirectoryAuditReport(
        input_dir=resolved_input_dir,
        sheet_name=sheet_name,
        generated_at=datetime.now(),
        source_audits=tuple(source_audits),
        missing_sources=tuple(missing_sources),
        source_issues=tuple(source_issues),
    )


def format_directory_audit_report(
    report: DirectoryAuditReport,
    *,
    log_path: Path | None = None,
) -> str:
    lines = [
        "客户映射核查结果",
        "",
        f"目录: {report.input_dir}",
        f"Sheet: {report.sheet_name}",
        f"检查时间: {report.generated_at.strftime('%Y-%m-%d %H:%M:%S')}",
        f"日志文件: {log_path.resolve() if log_path is not None else '未写入'}",
        "",
        "摘要:",
        f"- 规则涉及来源文件数: {report.expected_source_file_count}",
        f"- 实际检查文件数: {report.checked_source_file_count}",
        f"- 存在未映射记录的文件数: {report.files_with_unmapped_records}",
        f"- 未映射记录数: {report.total_unmapped_records}",
    ]

    lines.extend(["", "未映射记录:"])
    source_audits_with_records = [audit for audit in report.source_audits if audit.unmapped_records]
    if not source_audits_with_records:
        lines.append("- 未发现未映射记录。")
    else:
        for audit in source_audits_with_records:
            lines.append(
                f"- {audit.source_file_name} | 数据标签列 {audit.data_column} | 辅助标签列 {audit.auxiliary_column or '无'} | 数据行数 {audit.total_rows}"
            )
            for record in audit.unmapped_records:
                fragments = [f"行 {record.excel_row_number}"]
                if record.auxiliary_column is not None:
                    auxiliary_value = record.auxiliary_value if record.auxiliary_value else "<空>"
                    fragments.append(
                        f"辅助标签({record.auxiliary_column})={auxiliary_value}"
                    )
                fragments.append(f"数据标签({record.data_column})={record.data_value}")
                fragments.append(record.reason)
                lines.append(f"  {' | '.join(fragments)}")

    if report.missing_sources or report.source_issues:
        lines.extend(["", "附加提示:"])
        if report.missing_sources:
            lines.append(
                "- 缺少来源文件: "
                + "、".join(item.source_file_name for item in report.missing_sources)
            )
        for issue in report.source_issues:
            lines.append(f"- {issue.source_file_name}: {issue.message}")

    return "\n".join(lines)


def write_audit_log(report_text: str, log_path: Path) -> Path:
    resolved_log_path = log_path.resolve()
    resolved_log_path.parent.mkdir(parents=True, exist_ok=True)
    resolved_log_path.write_text(report_text + "\n", encoding="utf-8")
    return resolved_log_path


def build_default_log_path(log_dir: Path, *, now: datetime | None = None) -> Path:
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return log_dir.resolve() / f"unmapped_customer_records_{timestamp}.log"


def resolve_log_path(args: argparse.Namespace) -> Path:
    if getattr(args, "log_file", None) is not None:
        return Path(args.log_file).resolve()
    return build_default_log_path(Path(args.log_dir))


def run_audit_command(args: argparse.Namespace) -> tuple[DirectoryAuditReport, Path, str]:
    report = run_directory_audit(args.input_dir, sheet_name=args.sheet_name)
    log_path = resolve_log_path(args)
    report_text = format_directory_audit_report(report, log_path=log_path)
    write_audit_log(report_text, log_path)
    print(report_text)
    return report, log_path, report_text


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="检查指定目录下的原始问卷文件，找出未被客户类别映射规则覆盖的具体记录。"
    )
    parser.add_argument("--input-dir", type=Path, required=True, help="要扫描的原始问卷目录")
    parser.add_argument(
        "--sheet-name",
        default=DEFAULT_SHEET_NAME,
        help=f"要读取的 sheet 名，默认 {DEFAULT_SHEET_NAME}",
    )
    parser.add_argument(
        "--log-dir",
        type=Path,
        default=Path(DEFAULT_LOG_DIR),
        help=f"日志目录，默认 {DEFAULT_LOG_DIR}",
    )
    parser.add_argument(
        "--log-file",
        type=Path,
        help="指定日志文件路径；如果传入，则优先于 --log-dir",
    )
    return parser.parse_args()


def main() -> None:
    run_audit_command(parse_args())


if __name__ == "__main__":
    main()

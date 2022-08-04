from typing import Any, List, Tuple

import openpyxl


def format_row(headers, row, report_date_column, is_reopened) -> dict[str, Any]:
    return {
        "datetime": row[headers.index(report_date_column)].value,
        "case_number": row[headers.index("Case Number")].value,
        "case_owner": row[headers.index("Case Owner")].value,
        "status": row[headers.index("Status")].value,
        "is_open": bool(int(row[headers.index("Open")].value)),
        "is_closed": bool(int(row[headers.index("Closed")].value)),
        "was_escalated": bool(int(row[headers.index("Was Escalated")].value)),
        "is_reopened": is_reopened,
    }


def get_cases(
    report_date, report_file, report_date_column, is_reopened, file_display_name
) -> List[dict[str, Any]]:
    wb = openpyxl.load_workbook(report_file)
    ws = wb.active

    try:
        cases = []
        for i, row in enumerate(ws.rows):
            if not i:
                headers = [c.value for c in row]
                report_date_index = headers.index(report_date_column)
                continue

            if row[report_date_index].value.date() != report_date:
                continue

            cases.append(format_row(headers, row, report_date_column, is_reopened))

        return cases

    except ValueError as e:
        raise ValueError(
            f"Unable to parse columns for {file_display_name} file; is it in the right format?"
        ) from e


def keep_unique_case_by_newest_datetime(
    cases: List[dict[str, Any]]
) -> List[dict[str, Any]]:
    # sort by case number asc, datetime desc
    cases.sort(key=lambda x: x["datetime"], reverse=True)
    cases.sort(key=lambda x: x["case_number"])

    # remove duplicates
    case_numbers = set()
    unique_cases: List[dict[str, Any]] = []
    for case in cases:
        if case["case_number"] not in case_numbers:
            case_numbers.add(case["case_number"])
            unique_cases.append(case)

    return unique_cases


def get_child_case_count(
    report_file, case_numbers, child_case_threshold, whitespace_offset=1
) -> int:
    wb = openpyxl.load_workbook(report_file)
    ws = wb.active
    child_case_count: int = 0

    # only look at cases we're interested in
    case_number = None

    # find the "Subtotal" rows and sum them if they exceed the threshold
    for row in ws.rows:
        if str(row[whitespace_offset].value).isnumeric():
            case_number = row[whitespace_offset].value

        if case_number in case_numbers and row[whitespace_offset].value == "Subtotal":
            case_count = row[whitespace_offset + 2].value
            if case_count >= child_case_threshold:
                child_case_count += case_count

    return child_case_count


def get_closed_and_escalated_case_counts(cases) -> Tuple[int, int]:
    # filter to only cases that are not re-opened
    cases = [case for case in cases if not case["is_reopened"]]

    escalated_case_count = len([case for case in cases if case["was_escalated"]])
    closed_case_count = len(cases) - escalated_case_count

    return closed_case_count, escalated_case_count


def main(
    report_date,
    fcr_reopened_file,
    fcr_closed_file,
    fcr_parent_file,
    child_case_threshold,
    child_case_count_override=None,
) -> dict[str, float | int]:

    cases = get_cases(
        report_date,
        fcr_reopened_file,
        "Edit Date",
        is_reopened=True,
        file_display_name="Re-opened Report",
    )

    cases.extend(
        get_cases(
            report_date,
            fcr_closed_file,
            "Date/Time Opened",
            is_reopened=False,
            file_display_name="Closed Report",
        )
    )

    cases = keep_unique_case_by_newest_datetime(cases)

    if not cases:
        raise ValueError("No cases found for the given report date")

    child_case_count = (
        child_case_count_override
        if child_case_count_override is not None
        else get_child_case_count(
            fcr_parent_file,
            [case["case_number"] for case in cases],
            child_case_threshold,
        )
    )

    closed_case_count, escalated_case_count = get_closed_and_escalated_case_counts(
        cases
    )

    fcr = (closed_case_count - escalated_case_count - child_case_count) / sum(
        [closed_case_count, escalated_case_count, child_case_count]
    )

    return {
        "fcr": fcr,
        "closed_case_count": closed_case_count,
        "escalated_case_count": escalated_case_count,
        "child_case_count": child_case_count,
    }

import re
from pathlib import Path

import pandas as pd

from mars_group import members
from work_folders import sharepoint_folder
from otl import fy_start
from work_tools import read_excel

fy = fy_start.year - 2000
FORECAST_FILE = (sharepoint_folder / 'ASTeC-Management - Finance' / f'{fy}-{fy + 1}' /
               'Forecast Model' / f'ASTeC Forecast Model FY{fy}_{fy + 1}.xlsx')
SHEET_NAME = "ASTeC Staff"

# Differences below this level are treated as rounding noise.
TOLERANCE = 0.005


def base_project_code(value: object) -> str | None:
    """
    Extract the base project code from a spreadsheet value.

    Examples
    --------
    STGA02003 1-2      -> STGA02003
    STLA00037-147      -> STLA00037
    STLA00037 -151     -> STLA00037
    STKA00226 task 02  -> STKA00226
    STGA00038          -> STGA00038
    """
    if pd.isna(value):
        return None

    text = str(value).strip().upper()

    # Normal STFC-style project identifier: four letters followed by five digits.
    match = re.match(r"([A-Z]{4}\d{5})", text)

    if match:
        return match.group(1)

    # Keep special non-project categories if you ever want to compare them.
    # if text in {"NATIONAL LABS", "OVERTIME", "STRA00009"}:
    #     return text

    return None


def normalise_person_name(value: object) -> str:
    """
    Convert a spreadsheet name to a comparison-friendly form.

    Examples
    --------
    'Bainbridge, Doctor Alexander Robert' -> 'alexander bainbridge'
    'Shepherd, Mr Benjamin John Arthur'   -> 'benjamin shepherd'
    """
    if pd.isna(value):
        return ""

    text = " ".join(str(value).strip().split())

    if "," not in text:
        return text.casefold()

    surname, given_names = text.split(",", maxsplit=1)

    # Remove common titles from the start of the given-name section.
    given_names = re.sub(
        r"^\s*(mr|mrs|miss|ms|doctor|dr|professor|prof)\.?\s+",
        "",
        given_names,
        flags=re.IGNORECASE,
    )

    first_name = given_names.strip().split()[0]

    return f"{first_name} {surname.strip()}".casefold()


def normalise_local_name(name: str) -> str:
    """
    Convert a GroupMember name such as 'Alexander Bainbridge'
    to the same comparison form used for spreadsheet names.
    """
    parts = " ".join(name.strip().split()).split()

    if len(parts) < 2:
        return name.casefold()

    return f"{parts[0]} {parts[-1]}".casefold()


def find_person_column(spreadsheet_names: pd.Series, local_name: str) -> int:
    """
    Find the spreadsheet column for one GroupMember.
    """
    matches = spreadsheet_names[
        spreadsheet_names.astype(str).str.startswith(local_name)
    ]

    if len(matches) == 1:
        return int(matches.index[0])

    if len(matches) == 0:
        raise KeyError(
            f"No column found for {local_name!r}; "
        )

    raise KeyError(
        f"Multiple columns found for {local_name!r}: "
        f"{list(matches.index)}"
    )


def load_department_bookings(
    filename: Path = FORECAST_FILE,
) -> tuple[pd.DataFrame, pd.Series]:
    """
    Load the ASTeC Staff sheet.

    Returns
    -------
    data:
        Raw project rows. Columns remain numbered because the spreadsheet has
        group headings in row 1 and employee names in row 2.

    names:
        Series mapping column number to employee name.
    """
    raw = read_excel(
        filename,
        sheet_name=SHEET_NAME,
        header=None,
        engine="openpyxl",
    )

    # Excel row 2 contains staff names. With zero-based pandas indexing,
    # that is row index 1.
    names = raw.iloc[1]

    # Excel row 3 onwards contains project records.
    data = raw.iloc[2:].copy()

    data = data.rename(
        columns={
            0: "spreadsheet_code",
            1: "project_name",
        }
    )

    data["project"] = data["spreadsheet_code"].map(base_project_code)

    # Drop totals, check rows, blank rows and other non-project records.
    data = data[data["project"].notna()].copy()

    return data, names


def department_bookings_for_member(
    data: pd.DataFrame,
    member_column: int,
) -> pd.Series:
    """
    Return departmental annual FTE by base project for one person.

    Multiple task rows belonging to the same project are summed here.
    """
    values = pd.to_numeric(
        data[member_column],
        errors="coerce",
    ).fillna(0.0)

    result = (
        pd.DataFrame(
            {
                "project": data["project"],
                "department_fte": values,
            }
        )
        .groupby("project", as_index=True)["department_fte"]
        .sum()
    )

    # Remove zero-only projects to keep the comparison concise.
    return result[result.abs() > 1e-12].sort_index()


def local_bookings_for_member(member) -> pd.Series:
    """
    Return BookingPlan annual FTE by base project for one GroupMember.

    If the member has multiple entries under different tasks of the same
    project, they are aggregated together.
    """
    records = []

    for entry in member.booking_plan.entries:
        project = base_project_code(entry.code.project)

        if project is None:
            continue

        records.append(
            {
                "project": project,
                "local_fte": float(entry.annual_fte or 0.0),
            }
        )

    if not records:
        return pd.Series(dtype=float, name="local_fte")

    return (
        pd.DataFrame(records)
        .groupby("project", as_index=True)["local_fte"]
        .sum()
        .sort_index()
    )


def compare_member(
    member,
    data: pd.DataFrame,
    names: pd.Series,
    tolerance: float = TOLERANCE,
) -> pd.DataFrame:
    """
    Compare one person's local booking plan with the department spreadsheet.
    """
    member_column = find_person_column(names, member.formal_name())

    department = department_bookings_for_member(
        data,
        member_column,
    )

    local = local_bookings_for_member(member)

    comparison = pd.concat(
        [local, department],
        axis=1,
    ).fillna(0.0)

    comparison["difference"] = (
        comparison["local_fte"]
        - comparison["department_fte"]
    )

    comparison["abs_difference"] = comparison["difference"].abs()

    comparison["status"] = "MATCH"

    comparison.loc[
        comparison["abs_difference"] > tolerance,
        "status",
    ] = "DIFFERENT"

    comparison.loc[
        (comparison["local_fte"] > tolerance)
        & (comparison["department_fte"].abs() <= tolerance),
        "status",
    ] = "LOCAL ONLY"

    comparison.loc[
        (comparison["department_fte"] > tolerance)
        & (comparison["local_fte"].abs() <= tolerance),
        "status",
    ] = "DEPARTMENT ONLY"

    comparison.insert(0, "member", member.name)

    return comparison.reset_index(names="project")


def compare_all_members(
    filename: Path = FORECAST_FILE,
    tolerance: float = TOLERANCE,
) -> pd.DataFrame:
    """
    Compare all MaRS members against the departmental forecast.
    """
    data, names = load_department_bookings(filename)

    comparisons = []

    for member in members:
        try:
            result = compare_member(
                member,
                data,
                names,
                tolerance=tolerance,
            )
        except KeyError as error:
            print(f"WARNING: {error}")
            continue

        comparisons.append(result)

    if not comparisons:
        return pd.DataFrame()

    result = pd.concat(
        comparisons,
        ignore_index=True,
    )

    status_order = {
        "DIFFERENT": 0,
        "LOCAL ONLY": 1,
        "DEPARTMENT ONLY": 2,
        "MATCH": 3,
    }

    result["_status_order"] = result["status"].map(status_order)

    result = (
        result.sort_values(
            ["_status_order", "member", "abs_difference"],
            ascending=[True, True, False],
        )
        .drop(columns="_status_order")
        .reset_index(drop=True)
    )

    return result


def print_summary(comparison: pd.DataFrame) -> None:
    """
    Print mismatches grouped by member.
    """
    differences = comparison[
        comparison["status"] != "MATCH"
    ]

    if differences.empty:
        print("All bookings match within the specified tolerance.")
        return

    for member_name, rows in differences.groupby(
        "member",
        sort=False,
    ):
        print(f"\n{member_name}")
        print("-" * len(member_name))

        print(
            rows[
                [
                    "project",
                    "local_fte",
                    "department_fte",
                    "difference",
                    "status",
                ]
            ].to_string(
                index=False,
                formatters={
                    "local_fte": "{:.4f}".format,
                    "department_fte": "{:.4f}".format,
                    "difference": "{:+.4f}".format,
                },
            )
        )


if __name__ == "__main__":
    # data, names = load_department_bookings()
    # for member in members:
    #     member_column = find_person_column(names, member.formal_name())
    #     print(member.known_as, member_column)
# else:
    comparison = compare_all_members()

    print_summary(comparison)

    # Full comparison, including matches.
    comparison.to_excel(
        "mars_forecast_comparison.xlsx",
        index=False,
        engine="openpyxl",
    )

    # Concise CSV containing only discrepancies.
    comparison.loc[
        comparison["status"] != "MATCH"
    ].to_csv(        "mars_forecast_differences.csv",
        index=False,
    )

    print("\nWritten:")
    print("  mars_forecast_comparison.xlsx")
    print("  mars_forecast_differences.csv")
from __future__ import annotations
from dataclasses import dataclass
from typing import Iterable, Union
from pathlib import Path
from openpyxl import load_workbook
from tabulate import tabulate
from datetime import datetime
import calendar
import locale

locale.setlocale(locale.LC_ALL, '')

#TODO: don't hardcode header rows; search for them instead

summary_header_row_idx = 17
detail_header_row_idx = 17
date_cell = 'A3'

MARKDOWN_TABLE_FORMATS = ['github', 'pipe']


def remove_suffix_starting_with(string: str, substrings: Iterable[str]) -> str:
    for substring in substrings:
        if substring in string:
            idx = string.index(substring)
            string = string[:idx].strip()
            return string
    return string


def remove_substrings(string: str, substrings: Iterable[str]) -> str:
    for substring in substrings:
        if substring in string:
            string = string.replace(substring, '')
    return string


def find_expenses_by_category(summary: ProjectSummary, ws_detail, project_name_header: str, project_name: str) -> None:
    detail_expense_category_header = 'Expenditure Category Name'
    detail_expenses_header = 'Expenses'
    detail_budget_header = 'Budget'
    detail_header_row = ws_detail[detail_header_row_idx]
    detail_col_idxs = {}
    for i, cell in enumerate(detail_header_row):
        header = cell.value
        if header in [project_name_header, detail_expense_category_header, detail_expenses_header,
                      detail_budget_header]:
            detail_col_idxs[header] = i

    found_project_name = False
    for row in ws_detail.iter_rows(min_row=18):
        if row[detail_col_idxs[project_name_header]].value == project_name:
            found_project_name = True
            break
    if not found_project_name:
        raise ValueError(f"Couldn't find expenses for project {project_name}")

    for row in ws_detail.iter_rows(min_row=18):
        if row[detail_col_idxs[project_name_header]].value == project_name:
            category = row[detail_col_idxs[detail_expense_category_header]].value
            if 'Salaries and Wages' in category:
                summary.salary = row[detail_col_idxs[detail_expenses_header]].value
            elif 'Travel' in category:
                summary.travel = row[detail_col_idxs[detail_expenses_header]].value
            elif 'Supplies / Services / Other Expenses' in category:
                summary.supplies = row[detail_col_idxs[detail_expenses_header]].value
            elif 'Fringe Benefits' in category:
                summary.fringe = row[detail_col_idxs[detail_expenses_header]].value
            elif 'Fellowship & Scholarships' in category:
                summary.fellowship = row[detail_col_idxs[detail_expenses_header]].value
            elif 'Indirect Costs' in category:
                summary.indirect = row[detail_col_idxs[detail_expenses_header]].value
            else:
                print(f"Unknown category {category}; consider adding it to the list of categories?")


def extract_date(ws_summary) -> datetime.date:
    raw = ws_summary[date_cell].value
    raw = raw.replace('Report Run Date:', '').strip()
    date_and_time = datetime.strptime(raw, '%Y-%m-%d %I:%M:%S %p')
    return date_and_time


@dataclass
class Summary:
    project_summaries: list[ProjectSummary]
    date_and_time: datetime

    def year(self) -> int:
        return self.date_and_time.year

    def month(self) -> str:
        return calendar.month_name[self.date_and_time.month]

    def day(self) -> int:
        return self.date_and_time.day

    def date(self) -> datetime.date:
        return self.date_and_time.date()

    @staticmethod
    def from_file(fn: Union[Path, str], substrings_to_clean: Iterable[str] = (),
                  suffixes_to_clean: Iterable[str] = ()) -> Summary:
        """
        Read the Excel file named `fn` (alternately `fn` can be a pathlib.Path object)
        and return a list of summaries of projects in the file.
        """
        if isinstance(fn, Path):
            fn = str(fn.resolve())
        wb = load_workbook(filename=fn, read_only=True)
        ws_summary = wb['Summary']
        ws_detail = wb['Detail']

        date = extract_date(ws_summary)

        header_row = ws_summary[summary_header_row_idx]
        assert header_row[0].value == "Award Number"
        project_name_header = 'Project Name'
        budget_header = 'Budget'
        expenses_header = 'Expenses'
        balance_header = 'Budget Balance (Budget – (Comm & Exp))'
        col_names = [project_name_header, budget_header, expenses_header, balance_header]

        # used to distinguish CS-specific accounts that all have the same project name of
        # "David Doty ENGR COMPUTER SCIENCE PPM Only"
        # In these cases we grab the Task name and used that as the project name instead
        task_header = 'Task/Subtask Name'

        summary_col_idxs = {}
        for i, cell in enumerate(header_row):
            header = cell.value
            if header in col_names + [task_header]:
                summary_col_idxs[header] = i

        project_summaries = []
        for row in ws_summary.iter_rows(min_row=summary_header_row_idx + 1):
            project_name = row[summary_col_idxs[project_name_header]].value
            budget = row[summary_col_idxs[budget_header]].value
            expenses = row[summary_col_idxs[expenses_header]].value
            balance = row[summary_col_idxs[balance_header]].value
            if None in [project_name, budget, expenses]:
                # some rows are empty; skip them
                continue

            budget = float(budget)
            expenses = float(expenses)
            balance = float(balance)
            summary = ProjectSummary(
                project_name=project_name,
                balance=balance,
                budget=budget,
                expenses=expenses,
                salary=0,
                travel=0,
                supplies=0,
                fringe=0,
                fellowship=0,
                indirect=0,
            )
            find_expenses_by_category(summary, ws_detail, project_name_header, project_name)

            clean_project_name = project_name
            if 'PPM Only' in project_name:
                # replace with more specific task name
                clean_project_name = row[summary_col_idxs[task_header]].value

            clean_project_name = remove_suffix_starting_with(clean_project_name, suffixes_to_clean)
            clean_project_name = remove_substrings(clean_project_name, substrings_to_clean)
            summary.project_name = clean_project_name

            project_summaries.append(summary)

        return Summary(project_summaries, date)

    def __str__(self) -> str:
        return self.table()

    def __repr__(self) -> str:
        return self.table()

    def table(self, tablefmt: str = 'rounded_outline') -> str:
        table = []
        for project_summary in self.project_summaries:
            row = [
                project_summary.project_name,
                locale.currency(project_summary.expenses, grouping=True),
                locale.currency(project_summary.salary, grouping=True),
                locale.currency(project_summary.travel, grouping=True),
                locale.currency(project_summary.supplies, grouping=True),
                locale.currency(project_summary.fringe, grouping=True),
                locale.currency(project_summary.fellowship, grouping=True),
                locale.currency(project_summary.indirect, grouping=True),
                locale.currency(project_summary.balance, grouping=True),
                locale.currency(project_summary.budget, grouping=True),
            ]
            if tablefmt in MARKDOWN_TABLE_FORMATS:
                # escape $ so markdown does not interpret it as Latex
                for i in range(1, len(row)):
                    row[i] = row[i].replace('$', r'\$')
            table.append(row)

        new_headers = ['Project Name', 'Expenses', 'Salary', 'Travel', 'Supplies', 'Fringe',
                       'Fellowship', 'Indirect', 'Balance', 'Budget']
        table_tabulated = tabulate(table, headers=new_headers, tablefmt=tablefmt,
                                   colalign=(
                                       'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right', 'right',
                                       'right'))
        return table_tabulated

    def diff_table(self, summary_earlier: Summary, tablefmt: str = 'rounded_outline') -> str:
        table = []
        for (summary_later, summary_earlier) in zip(self.project_summaries, summary_earlier.project_summaries):
            if summary_later.project_name != summary_earlier.project_name:
                raise ValueError("Can't diff summaries with different project names")
            diff = summary_later.diff(summary_earlier)
            row = [
                diff.project_name,
                locale.currency(diff.expenses, grouping=True),
                locale.currency(diff.salary, grouping=True),
                locale.currency(diff.travel, grouping=True),
                locale.currency(diff.supplies, grouping=True),
                locale.currency(diff.fringe, grouping=True),
                locale.currency(diff.fellowship, grouping=True),
                locale.currency(diff.indirect, grouping=True),
                locale.currency(diff.balance, grouping=True),
            ]
            if tablefmt in MARKDOWN_TABLE_FORMATS:
                # escape $ so markdown does not interpret it as Latex
                for i in range(1, len(row)):
                    row[i] = row[i].replace('$', r'\$')
            table.append(row)

        new_headers = ['Project Name', 'Expenses', 'Salary', 'Travel', 'Supplies', 'Fringe',
                       'Fellowship', 'Indirect', 'Balance']
        table_tabulated = tabulate(table, headers=new_headers, tablefmt=tablefmt,
                                   colalign=(
                                       'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right', 'right'))
        return table_tabulated


@dataclass
class ProjectSummary:
    project_name: str
    balance: float
    budget: float
    expenses: float
    salary: float
    travel: float
    supplies: float
    fringe: float
    fellowship: float
    indirect: float

    def diff(self, other: ProjectSummary) -> ProjectSummary:
        if self.project_name != other.project_name:
            raise ValueError("Can't diff summaries with different project names")
        return ProjectSummary(
            project_name=self.project_name,
            balance=self.balance - other.balance,
            budget=0,
            expenses=self.expenses - other.expenses,
            salary=self.salary - other.salary,
            travel=self.travel - other.travel,
            supplies=self.supplies - other.supplies,
            fringe=self.fringe - other.fringe,
            fellowship=self.fellowship - other.fellowship,
            indirect=self.indirect - other.indirect,
        )


if __name__ == '__main__':
    suffixes = ['K3023', 'DOTY DEFAULT PROJECT 13U00']
    substrings = ['Doty', 'CS ']
    summary_aug = Summary.from_file('2024-8-5.xlsx', substrings_to_clean=substrings, suffixes_to_clean=suffixes)
    summary_jul = Summary.from_file('2024-7-11.xlsx', substrings_to_clean=substrings, suffixes_to_clean=suffixes)

    print(f"Totals for {summary_aug.month()} {summary_aug.year()}")
    print(summary_aug.table())
    print(f"Totals for {summary_jul.month()} {summary_jul.year()}")
    print(summary_jul.table())
    print(f"Difference between {summary_aug.month()} and {summary_jul.month()}")
    print(summary_aug.diff_table(summary_jul))

from __future__ import annotations
from dataclasses import dataclass
from typing import Iterable, Union, List, Optional
from pathlib import Path
from openpyxl import load_workbook
from tabulate import tabulate
from datetime import datetime
import calendar


def format_currency(amount: float) -> str:
    # The "proper" way to do this is with the locale module,
    # but you need the system to have certain locales installed,
    # and I don't want users to run into those stupid errors,
    # so we just manually format the currency.
    return f"${amount:,.2f}" if amount >= 0 else f"-${abs(amount):,.2f}"


# TODO: don't hardcode header rows; search for them instead

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
    # make a defensive copy since we modify the list
    substrings = list(substrings)

    # sort substrings by substring relation to each other to avoid removing 
    # a subsubstring that causes the remaining part not to be removed
    # e.g., if we have "abc" and "a", we want to remove "abc" first, otherwise
    # if string is "a abc My Project", we will end up changing like this:
    # " bc My Project" instead of " My Project" because removing the 
    # subsubstring "a" first changes the other substring "abc" to "bc"
    for i, sub1 in enumerate(substrings):
        for j in range(i + 1, len(substrings)):
            sub2 = substrings[j]
            if sub1 in sub2:
                substrings[i], substrings[j] = sub2, sub1

    for substring in substrings:
        if substring in string:
            string = string.replace(substring, '')
    return string


def clean_whitespace(string: str) -> str:
    return ' '.join(string.split())


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


POSSIBLE_HEADERS = ['Expenses', 'Salary', 'Travel', 'Supplies', 'Fringe', 'Fellowship', 'Indirect', 'Balance', 'Budget']


@dataclass
class Summary:
    project_summaries: List[ProjectSummary]
    date_and_time: datetime
    headers: List[str]

    def year(self) -> int:
        """The year of the summary (as an integer)"""
        return self.date_and_time.year

    def month(self) -> str:
        """The month of the summary (as a string)"""
        return calendar.month_name[self.date_and_time.month]

    def day(self) -> int:
        """The day of the summary (as an integer)"""
        return self.date_and_time.day

    def date(self) -> datetime.date:
        """The date of the summary (as a datetime.date object)"""
        return self.date_and_time.date()

    @staticmethod
    def from_file(
            fn: Union[Path, str],
            substrings_to_clean: Iterable[str] = (),
            suffixes_to_clean: Iterable[str] = (),
            headers: Optional[Iterable[str]] = None,
    ) -> Summary:
        """
        Read the Excel file named `fn` (alternately `fn` can be a pathlib.Path object)
        and return a list of summaries of projects in the file.

        Args:
            fn: The filename (or pathlib.Path object) of the AggieExpense Excel file to read
            substrings_to_clean: A list of substrings to remove from the project name
            suffixes_to_clean: A list of substrings to remove from the project name, including
                the whole suffix following the substring
            headers: A list of headers to include in the summary; must be a subset of
                ['Expenses', 'Salary', 'Travel', 'Supplies', 'Fringe', 'Fellowship', 'Indirect', 'Balance', 'Budget']
                The headers will be displayed in the order they are given (so is also a way to reorder them
                from the default order even if you include all of them)
                If not specified they are all displayed in that order for both the `table` and `diff_table` methods,
                although the `diff_table` method will not display the 'Budget' column since it should never change
                in principle.
        Returns:
            A Summary object containing the summaries of all the projects in the file.
        """
        if headers is None:
            headers = POSSIBLE_HEADERS
        for header in headers:
            if header not in POSSIBLE_HEADERS:
                raise ValueError(f"Invalid heading: {header}; must be one of {', '.join(POSSIBLE_HEADERS)}")

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
            # TODO: replace this with a check for Project Type = Internal (vs. Sponsored)
            if 'PPM Only' in project_name:
                # replace with more specific task name
                clean_project_name = row[summary_col_idxs[task_header]].value

            clean_project_name = remove_suffix_starting_with(clean_project_name, suffixes_to_clean)
            clean_project_name = remove_substrings(clean_project_name, substrings_to_clean)
            clean_project_name = clean_whitespace(clean_project_name)
            summary.project_name = clean_project_name

            project_summaries.append(summary)

        return Summary(project_summaries, date, headers)

    def __str__(self) -> str:
        return self.table()

    def __repr__(self) -> str:
        return self.table()

    def table(self, tablefmt: str = 'rounded_outline') -> str:
        """
        Return a string representation of the summary as a string in tabular form.

        Args:
            tablefmt: The format of the table; see the Python package tabulate documentation for options

        Returns:
            A string representation of the summary as a string in tabular form
        """
        table = []
        for project_summary in self.project_summaries:
            header_to_field = {
                'Expenses': project_summary.expenses,
                'Salary': project_summary.salary,
                'Travel': project_summary.travel,
                'Supplies': project_summary.supplies,
                'Fringe': project_summary.fringe,
                'Fellowship': project_summary.fellowship,
                'Indirect': project_summary.indirect,
                'Balance': project_summary.balance,
                'Budget': project_summary.budget,
            }
            row = [project_summary.project_name]
            for header in self.headers:
                row.append(format_currency(header_to_field[header]))

            if tablefmt in MARKDOWN_TABLE_FORMATS:
                # escape $ so markdown does not interpret it as Latex
                for i in range(1, len(row)):
                    row[i] = row[i].replace('$', r'\$')
            table.append(row)

        new_headers = ['Project Name'] + self.headers
                       # , 'Expenses', 'Salary', 'Travel', 'Supplies', 'Fringe',)
                       # 'Fellowship', 'Indirect', 'Balance', 'Budget']
        colalign = ['left'] + ['right'] * len(self.headers)
        table_tabulated = tabulate(table, headers=new_headers, tablefmt=tablefmt, colalign=colalign)
        return table_tabulated

    def diff_table(self, summary_earlier: Summary, tablefmt: str = 'rounded_outline') -> str:
        """
        Return a string representation of the differences between this summary and `summary_earlier`.

        Args:
            tablefmt: The format of the table; see the Python package tabulate documentation for options

        Returns:
            A string representation of the summary of differences as a string in tabular form
        """
        table = []
        for (summary_later, summary_earlier) in zip(self.project_summaries, summary_earlier.project_summaries):
            if summary_later.project_name != summary_earlier.project_name:
                raise ValueError("Can't diff summaries with different project names")
            diff = summary_later.diff(summary_earlier)
            header_to_field = {
                'Expenses': diff.expenses,
                'Salary': diff.salary,
                'Travel': diff.travel,
                'Supplies': diff.supplies,
                'Fringe': diff.fringe,
                'Fellowship': diff.fellowship,
                'Indirect': diff.indirect,
                'Balance': diff.balance,
            }
            row = [diff.project_name]
            for header in self.headers:
                if header == 'Budget': # don't show budget diff since it's always equal between two summaries
                    continue
                row.append(format_currency(header_to_field[header]))

            if tablefmt in MARKDOWN_TABLE_FORMATS:
                # escape $ so markdown does not interpret it as Latex
                for i in range(1, len(row)):
                    row[i] = row[i].replace('$', r'\$')
            table.append(row)

        new_headers = ['Project Name'] + self.headers
        num_expense_cols = len(self.headers) - 1 if 'Budget' in self.headers else len(self.headers)
        colalign = ['left'] + ['right'] * num_expense_cols
        table_tabulated = tabulate(table, headers=new_headers, tablefmt=tablefmt, colalign=colalign)
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

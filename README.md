# aggie-unterprise


## Overview
Here's an example of a useful summary of grant data:

Totals for August
| Project Name                       |      Balance |     Expenses |       Salary |      Travel |   Supplies |      Fringe |   Fellowship |     Indirect |       Budget |
|------------------------------------|--------------|--------------|--------------|-------------|------------|-------------|--------------|--------------|--------------|
| INDIRECT COST RETURN               |     \$904.00 |       \$0.00 |       \$0.00 |      \$0.00 |     \$0.00 |      \$0.00 |       \$0.00 |       \$0.00 |     \$904.00 |
| DISCRETIONARY FUNDS                |   \$2,500.00 |       \$0.00 |       \$0.00 |      \$0.00 |     \$0.00 |      \$0.00 |       \$0.00 |       \$0.00 |   \$2,500.00 |
| NSF Engineering DNA and RNA        | \$318,683.39 |  \$61,316.61 |  \$34,800.00 |      \$0.00 |   \$133.40 |  \$3,263.70 |       \$0.00 |  \$23,119.51 | \$380,000.00 |
| NSF CAREER Chemical Computation    |  \$17,180.28 | \$468,000.72 | \$211,746.21 | \$33,334.25 | \$8,847.54 | \$58,519.35 |   \$5,166.81 | \$150,386.56 | \$485,181.00 |
| REU CAREER Chemical Computation    |  \$18,750.37 |  \$44,062.63 |  \$43,180.29 |      \$0.00 |     \$0.00 |    \$882.34 |       \$0.00 |       \$0.00 |  \$62,813.00 |
| DOE Office of Science Basic Energy |  \$51,372.51 |  \$15,045.49 |   \$8,642.86 |      \$0.00 |     \$0.00 |    \$760.57 |       \$0.00 |   \$5,642.06 |  \$66,418.00 |

Totals for July
| Project Name                       |       Balance |     Expenses |       Salary |      Travel |   Supplies |      Fringe |   Fellowship |     Indirect |       Budget |
|------------------------------------|---------------|--------------|--------------|-------------|------------|-------------|--------------|--------------|--------------|
| INDIRECT COST RETURN               |      \$904.00 |       \$0.00 |       \$0.00 |      \$0.00 |     \$0.00 |      \$0.00 |       \$0.00 |       \$0.00 |     \$904.00 |
| DISCRETIONARY FUNDS                |    \$2,500.00 |       \$0.00 |       \$0.00 |      \$0.00 |     \$0.00 |      \$0.00 |       \$0.00 |       \$0.00 |   \$2,500.00 |
| NSF Engineering DNA and RNA        |  \$351,084.80 |  \$28,915.20 |  \$16,500.00 |      \$0.00 |   \$120.00 |  \$1,452.00 |       \$0.00 |  \$10,843.20 | \$380,000.00 |
| NSF CAREER Chemical Computation    | (\$44,997.85) | \$453,652.85 | \$206,470.48 | \$29,829.98 | \$6,388.58 | \$58,419.11 |   \$8,993.12 | \$143,551.58 | \$408,655.00 |
| REU CAREER Chemical Computation    |   \$18,750.37 |  \$44,062.63 |  \$43,180.29 |      \$0.00 |     \$0.00 |    \$882.34 |       \$0.00 |       \$0.00 |  \$62,813.00 |
| DOE Office of Science Basic Energy |   \$51,372.51 |  \$15,045.49 |   \$8,642.86 |      \$0.00 |     \$0.00 |    \$760.57 |       \$0.00 |   \$5,642.06 |  \$66,418.00 |

Difference between August and July
| Project Name                       |    Expenses |      Salary |     Travel |   Supplies |     Fringe |   Fellowship |    Indirect |       Balance |
|------------------------------------|-------------|-------------|------------|------------|------------|--------------|-------------|---------------|
| INDIRECT COST RETURN               |      \$0.00 |      \$0.00 |     \$0.00 |     \$0.00 |     \$0.00 |       \$0.00 |      \$0.00 |        \$0.00 |
| DISCRETIONARY FUNDS                |      \$0.00 |      \$0.00 |     \$0.00 |     \$0.00 |     \$0.00 |       \$0.00 |      \$0.00 |        \$0.00 |
| NSF Engineering DNA and RNA        | \$32,401.41 | \$18,300.00 |     \$0.00 |    \$13.40 | \$1,811.70 |       \$0.00 | \$12,276.31 | (\$32,401.41) |
| NSF CAREER Chemical Computation    | \$14,347.87 |  \$5,275.73 | \$3,504.27 | \$2,458.96 |   \$100.24 | (\$3,826.31) |  \$6,834.98 |   \$62,178.13 |
| REU CAREER Chemical Computation    |      \$0.00 |      \$0.00 |     \$0.00 |     \$0.00 |     \$0.00 |       \$0.00 |      \$0.00 |        \$0.00 |
| DOE Office of Science Basic Energy |      \$0.00 |      \$0.00 |     \$0.00 |     \$0.00 |     \$0.00 |       \$0.00 |      \$0.00 |        \$0.00 |

The first two tables summarize total balance, expenses, broken down by type of expenses, for several grants for two different months. These are the totals since the start of the grant. Since we sometimes care about monthly spending, the third table shows the *differences* between months, to indicate for instance, how much money was spent on supplies during July ($2,458.96 spent during July = $8,847.54 total spent by August - $6,388.58 total spent by July).

[AggieEnterprise](https://aggieenterprise.ucdavis.edu/) is a software system used by [UC Davis](https://www.ucdavis.edu/), whose purpose is to take the useful information above and bury it underneath loads of gibberish, resulting in a spreadsheet with redundant and useless information like this:

![AggieEnterprise spreadsheet screenshot](images/spreadsheet.png)

The aggie_unterprise package helps you, the **Aggie**, to **un**do this en**terprising** effort and view only the important data related to your grants.


## Example of usage
Suppose you have generated two spreadsheets from AggieEnterprise from two different months, named `2024-7-1.xlsx` and `2024-8-1.xlsx`. Then the following code will print text tables like the above:

```python
from aggie_unterprise import Summary, summary_diff_table, summary_table

summary_aug = Summary.from_file('2024-8-1.xlsx')
summary_jul = Summary.from_file('2024-7-1.xlsx')
table_diff = summary_diff_table(summary_aug, summary_jul)
table_aug = summary_table(summary_aug)
table_jul = summary_table(summary_jul)

print("Totals for August")
print(table_aug)
print("Totals for July")
print(table_jul)
print("Difference between August and July")
print(table_diff)
```

This will print a text table in a format similar this:

```
Totals for August
╭────────────────────────────────────┬─────────────┬─────────────┬─────────────┬────────────┬────────────┬────────────┬──────────────┬─────────────┬─────────────╮
│ Project Name                       │     Balance │    Expenses │      Salary │     Travel │   Supplies │     Fringe │   Fellowship │    Indirect │      Budget │
├────────────────────────────────────┼─────────────┼─────────────┼─────────────┼────────────┼────────────┼────────────┼──────────────┼─────────────┼─────────────┤
│ INDIRECT COST RETURN               │     $904.00 │       $0.00 │       $0.00 │      $0.00 │      $0.00 │      $0.00 │        $0.00 │       $0.00 │     $904.00 │
│ DISCRETIONARY FUNDS                │   $2,500.00 │       $0.00 │       $0.00 │      $0.00 │      $0.00 │      $0.00 │        $0.00 │       $0.00 │   $2,500.00 │
│ NSF Engineering DNA and RNA        │ $318,683.39 │  $61,316.61 │  $34,800.00 │      $0.00 │    $133.40 │  $3,263.70 │        $0.00 │  $23,119.51 │ $380,000.00 │
│ NSF CAREER Chemical Computation    │  $17,180.28 │ $468,000.72 │ $211,746.21 │ $33,334.25 │  $8,847.54 │ $58,519.35 │    $5,166.81 │ $150,386.56 │ $485,181.00 │
│ REU CAREER Chemical Computation    │  $18,750.37 │  $44,062.63 │  $43,180.29 │      $0.00 │      $0.00 │    $882.34 │        $0.00 │       $0.00 │  $62,813.00 │
│ DOE Office of Science Basic Energy │  $51,372.51 │  $15,045.49 │   $8,642.86 │      $0.00 │      $0.00 │    $760.57 │        $0.00 │   $5,642.06 │  $66,418.00 │
╰────────────────────────────────────┴─────────────┴─────────────┴─────────────┴────────────┴────────────┴────────────┴──────────────┴─────────────┴─────────────╯
Totals for July
╭────────────────────────────────────┬──────────────┬─────────────┬─────────────┬────────────┬────────────┬────────────┬──────────────┬─────────────┬─────────────╮
│ Project Name                       │      Balance │    Expenses │      Salary │     Travel │   Supplies │     Fringe │   Fellowship │    Indirect │      Budget │
├────────────────────────────────────┼──────────────┼─────────────┼─────────────┼────────────┼────────────┼────────────┼──────────────┼─────────────┼─────────────┤
│ INDIRECT COST RETURN               │      $904.00 │       $0.00 │       $0.00 │      $0.00 │      $0.00 │      $0.00 │        $0.00 │       $0.00 │     $904.00 │
│ DISCRETIONARY FUNDS                │    $2,500.00 │       $0.00 │       $0.00 │      $0.00 │      $0.00 │      $0.00 │        $0.00 │       $0.00 │   $2,500.00 │
│ NSF Engineering DNA and RNA        │  $351,084.80 │  $28,915.20 │  $16,500.00 │      $0.00 │    $120.00 │  $1,452.00 │        $0.00 │  $10,843.20 │ $380,000.00 │
│ NSF CAREER Chemical Computation    │ ($44,997.85) │ $453,652.85 │ $206,470.48 │ $29,829.98 │  $6,388.58 │ $58,419.11 │    $8,993.12 │ $143,551.58 │ $408,655.00 │
│ REU CAREER Chemical Computation    │   $18,750.37 │  $44,062.63 │  $43,180.29 │      $0.00 │      $0.00 │    $882.34 │        $0.00 │       $0.00 │  $62,813.00 │
│ DOE Office of Science Basic Energy │   $51,372.51 │  $15,045.49 │   $8,642.86 │      $0.00 │      $0.00 │    $760.57 │        $0.00 │   $5,642.06 │  $66,418.00 │
╰────────────────────────────────────┴──────────────┴─────────────┴─────────────┴────────────┴────────────┴────────────┴──────────────┴─────────────┴─────────────╯
Difference between August and July
╭────────────────────────────────────┬────────────┬────────────┬───────────┬────────────┬───────────┬──────────────┬────────────┬──────────────╮
│ Project Name                       │   Expenses │     Salary │    Travel │   Supplies │    Fringe │   Fellowship │   Indirect │      Balance │
├────────────────────────────────────┼────────────┼────────────┼───────────┼────────────┼───────────┼──────────────┼────────────┼──────────────┤
│ INDIRECT COST RETURN               │      $0.00 │      $0.00 │     $0.00 │      $0.00 │     $0.00 │        $0.00 │      $0.00 │        $0.00 │
│ DISCRETIONARY FUNDS                │      $0.00 │      $0.00 │     $0.00 │      $0.00 │     $0.00 │        $0.00 │      $0.00 │        $0.00 │
│ NSF Engineering DNA and RNA        │ $32,401.41 │ $18,300.00 │     $0.00 │     $13.40 │ $1,811.70 │        $0.00 │ $12,276.31 │ ($32,401.41) │
│ NSF CAREER Chemical Computation    │ $14,347.87 │  $5,275.73 │ $3,504.27 │  $2,458.96 │   $100.24 │  ($3,826.31) │  $6,834.98 │   $62,178.13 │
│ REU CAREER Chemical Computation    │      $0.00 │      $0.00 │     $0.00 │      $0.00 │     $0.00 │        $0.00 │      $0.00 │        $0.00 │
│ DOE Office of Science Basic Energy │      $0.00 │      $0.00 │     $0.00 │      $0.00 │     $0.00 │        $0.00 │      $0.00 │        $0.00 │
╰────────────────────────────────────┴────────────┴────────────┴───────────┴────────────┴───────────┴──────────────┴────────────┴──────────────╯
```

In the final diff table, one would normal expect each entry under Balance (which represents a change in balance from July to August) to be the negative of the entry under Expenses (total amount of expenses between July and August), as in the project "NSF Engineering DNA and RNA". However, sometimes a grant agency will deposit new funds (as happened in the "NSF CAREER Chemical Computation" entry above), so not always.

You can also render the tables in Markdown in a Jupyter notebook, so they will appear similar to the first tables shown above:

```python
from aggie_unterprise import Summary, summary_diff_table, summary_table

summary_aug = Summary.from_file('2024-8-5.xlsx')
summary_jul = Summary.from_file('2024-7-11.xlsx')
table_diff = summary_diff_table(summary_aug, summary_jul, tablefmt='github')
table_aug = summary_table(summary_aug, tablefmt='github')
table_jul = summary_table(summary_jul, tablefmt='github')

from IPython.display import display, Markdown
print("Totals for August")
display(Markdown(table_aug))
print("Totals for July")
display(Markdown(table_jul))
print("Difference between August and July")
display(Markdown(table_diff))
```

Finally, you can customize a bit how to clean up project names. They are taken either from the column named "Project Name" in the spreadsheet, unless that name has the substring `"PPM Only"` in it, which generally appear in department-specific funds like startup grants or indirect cost return, and are identical (for the CS department for me, they are all named "David Doty ENGR COMPUTER SCIENCE PPM Only"). For these funds we instead use the column "Task/Subtask Name" (which is useless for normal grants since it just says "TASK01", but is a bit more informative for department-specific funds, such as "INDIRECT COST RETURN" and "DISCRETIONARY FUNDS" above).

To clean up the project names, you can specify two parameters to the `Summary.from_file` method: `substrings_to_clean` and `suffixes_to_clean`. Any substring appearing in `substrings_to_clean` will be removed, for example if I set `substrings_to_clean=['CS', 'Doty']`, it will change the project name `"CS NSF Engineering DNA and RNA Doty K302325F33"` to `"NSF Engineering DNA and RNA  K302325F33"`. Anything in `suffixes_to_clean` will be removed, not only that substring, but the entire rest of the name. For instance, if I set `suffixes_to_clean=['K3023']`, it will change `"NSF Engineering DNA and RNA  K302325F33"` to `"NSF Engineering DNA and RNA"`.

## Installation
I may put this on [PyPI](https://pypi.org/) eventually so that it can be installed via pip. Until then you have to install the hard way:

1. **Clone the repo**: `git clone https://github.com/dave-doty/aggie-unterprise.git`

2. **Add to PYTHONPATH**: Assuming for example that you cloned the repository to the directory `C:\git\aggie-enterprise`, add the directory to your PYTHONPATH. In Windows this is done by going to settings and searching for "Environment Variables":\
![](images/env-var-search.png)\
and editing or adding (if necessary) a variable named PYTHONPATH with value `C:\git\aggie-enterprise`:\
![](images/env-var-set.png)\
In Linux/Mac, using the bash shell, this can be done by adding the line `PYTHONPATH=$PYTHONPATH:/mnt/c/git/aggie-enterprise` to the file `.bashrc` in your home directory.

3. **Install dependencies**: Type `pip install openpyxl tabulate` at the command line.

4. **Test**: Open a Python interpreter or Jupyter notebook and type `import aggie_unterprise`; it should import without errors.
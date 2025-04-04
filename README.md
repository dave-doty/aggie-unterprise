# aggie-unterprise

This document is intended to be read on [github](https://github.com/dave-doty/aggie-unterprise#readme). Some of what appears below does not render properly on other websites such as pypi.org.

## Table of contents

* [Installation](#installation)
* [Overview](#overview)
* [Examples](#examples)
* [API](#api)
* [Standalone program](#standalone-program)


## Installation

This assumes you have [Python](https://www.python.org/) installed (and [pip](https://pip.pypa.io/en/stable/installation/), which comes with Python unless you have a very old Python distribution). At a command prompt (e.g., cmd or Powershell in Windows, terminal in Linux or Mac), type 
```
pip install aggie-unterprise
````
If you don't want to install Python or write Python code, some functionality is available through a downloadable [standalone program](#standalone-program).


## Overview
This is an example of a useful summary of research grant funds:

### Totals for August 2024
| Project Name                       |  Expenses |    Salary |   Travel |  Supplies |   Fringe |  Fellowship |  Indirect |   Balance |    Budget |
|------------------------------------|-----------|-----------|----------|-----------|----------|-------------|-----------|-----------|-----------|
| INDIRECT COST RETURN               |           |           |          |           |          |             |           |     \$904 |     \$904 |
| DISCRETIONARY FUNDS                |           |           |          |           |          |             |           |   \$2,500 |   \$2,500 |
| NSF Engineering DNA and RNA        |  \$61,316 |  \$34,800 |          |     \$133 |  \$3,263 |             |  \$23,119 | \$318,683 | \$380,000 |
| NSF CAREER Chemical Computation    | \$468,000 | \$211,746 | \$33,334 |   \$8,847 | \$58,519 |     \$5,166 | \$150,386 |  \$17,180 | \$485,181 |
| REU CAREER Chemical Computation    |  \$44,062 |  \$43,180 |          |           |    \$882 |             |           |  \$18,750 |  \$62,813 |
| DOE Office of Science Basic Energy |  \$15,045 |   \$8,642 |          |           |    \$760 |             |   \$5,642 |  \$51,372 |  \$66,418 |

[AggieEnterprise](https://aggieenterprise.ucdavis.edu/) is a software system used by [UC Davis](https://www.ucdavis.edu/), whose purpose is bury this useful information beneath mountains of gibberish, resulting in a spreadsheet filled with useless trash like this:

![AggieEnterprise spreadsheet screenshot](images/spreadsheet.png)

The aggie_unterprise Python package helps you, the **AGGIE**, to **UN**do this en**TERPRIS**ing feat. It sifts through the trash to find the important data related to your grants.


## Examples
Suppose you have generated a spreadsheet from AggieEnterprise named `2024-8-1.xlsx` following [these instructions](https://servicehub.ucdavis.edu/servicehub?id=ucd_kb_article&sys_id=cc1942f61b32c6d80e0b2068b04bcbbf). The following code:

```python
from aggie_unterprise import Summary
summary = Summary.from_file('2024-8-1.xlsx')
print(f'Totals for {summary.month()} {summary.year()}')
print(summary)
```

will print something like this:

```
Totals for August 2024
╭────────────────────────────────────┬──────────┬──────────┬─────────┬──────────┬─────────┬────────────┬──────────┬──────────┬──────────╮
│ Project Name                       │ Expenses │   Salary │  Travel │ Supplies │  Fringe │ Fellowship │ Indirect │  Balance │   Budget │
├────────────────────────────────────┼──────────┼──────────┼─────────┼──────────┼─────────┼────────────┼──────────┼──────────┼──────────┤
│ INDIRECT COST RETURN               │          │          │         │          │         │            │          │     $904 │     $904 │
│ DISCRETIONARY FUNDS                │          │          │         │          │         │            │          │   $2,500 │   $2,500 │
│ NSF Engineering DNA and RNA        │  $61,316 │  $34,800 │         │     $133 │  $3,263 │            │  $23,119 │ $318,683 │ $380,000 │
│ NSF CAREER Chemical Computation    │ $468,000 │ $211,746 │ $33,334 │   $8,847 │ $58,519 │     $5,166 │ $150,386 │  $17,180 │ $485,181 │
│ REU CAREER Chemical Computation    │  $44,062 │  $43,180 │         │          │    $882 │            │          │  $18,750 │  $62,813 │
│ DOE Office of Science Basic Energy │  $15,045 │   $8,642 │         │          │    $760 │            │   $5,642 │  $51,372 │  $66,418 │
╰────────────────────────────────────┴──────────┴──────────┴─────────┴──────────┴─────────┴────────────┴──────────┴──────────┴──────────╯
```

The table summarizes expenses, broken down by type of expense, remaining balance, and original total budget, for each grant. These are the totals since the start of each grant. 

Since we sometimes care about monthly spending, we may want to know the *differences* between months, to indicate for instance, how much money was spent during July (e.g., $800 total spent by August - $600 total spent by July = $200 spent *during* July). The method `Summary.diff_table` gives this information:

```python
from aggie_unterprise import Summary
summary_aug = Summary.from_file('2024-8-1.xlsx')
summary_jul = Summary.from_file('2024-7-1.xlsx')
print(f'Difference between {summary_aug.month()} and {summary_jul.month()}')
print(summary_aug.diff_table(summary_jul))
```

```
Difference between August and July
╭────────────────────────────────────┬──────────┬─────────┬────────┬──────────┬────────┬────────────┬──────────┬──────────╮
│ Project Name                       │ Expenses │  Salary │ Travel │ Supplies │ Fringe │ Fellowship │ Indirect │  Balance │
├────────────────────────────────────┼──────────┼─────────┼────────┼──────────┼────────┼────────────┼──────────┼──────────┤
│ INDIRECT COST RETURN               │          │         │        │          │        │            │          │          │
│ DISCRETIONARY FUNDS                │          │         │        │          │        │            │          │          │
│ NSF Engineering DNA and RNA        │  $32,401 │ $18,300 │        │      $13 │ $1,811 │            │  $12,276 │ -$32,401 │
│ NSF CAREER Chemical Computation    │  $14,347 │  $5,275 │ $3,504 │   $2,458 │   $100 │    -$3,826 │   $6,834 │  $62,178 │
│ REU CAREER Chemical Computation    │          │         │        │          │        │            │          │          │
│ DOE Office of Science Basic Energy │          │         │        │          │        │            │          │          │
╰────────────────────────────────────┴──────────┴─────────┴────────┴──────────┴────────┴────────────┴──────────┴──────────╯
```

In the diff table, one would normally expect each entry under Balance (which represents a change in balance from July to August) to be the negative of the entry under Expenses (total amount of expenses between July and August), as in the project "NSF Engineering DNA and RNA". However, sometimes a grant agency will deposit new funds in between the dates (as happened in the "NSF CAREER Chemical Computation" entry above), so the change in balance and change in expenses are not always negatives of each other.

You can also render the tables in Markdown in a Jupyter notebook, so they will appear similar to the first table shown at the top of this document. When you print/stringify a `Summary`, it calls a method called `table` (so `f'{summary}'` is equivalent to `f'{summary.table()}'`), which, along with the method `diff_table`, can take the `tablefmt` argument with value `'github'` to render the table appropriately for Markdown (`'github'` means ["GitHub-flavored Markdown"](https://github.github.com/gfm/), which [Jupyter](https://jupyter.org/) can render nicely):

```python
from aggie_unterprise import Summary, summary_diff_table, summary_table

summary_aug = Summary.from_file('2024-8-5.xlsx')
summary_jul = Summary.from_file('2024-7-11.xlsx')

from IPython.display import display, Markdown
display(Markdown(f"""\
### Totals for {summary_aug.month()}
{summary_aug.table(tablefmt='github')}
### Totals for {summary_jul.month()}
{summary_jul.table(tablefmt='github')}
### Difference between {summary_aug.month()} and {summary_jul.month()}
{summary_aug.diff_table(summary_jul, tablefmt='github')}
"""))
```

In general, you can pass any value to `tablefmt` that the [tabulate](https://pypi.org/project/tabulate/) package expects in its function tabulate: https://github.com/astanin/python-tabulate?tab=readme-ov-file#table-format.

Finally, you can customize a bit how to clean up project names. They are taken either from the column "Project Name" in the spreadsheet for sponsored (external) grant, or the column "Task/Subtask Name" for internal funds (e.g., startup grants or indirect cost return), since the project names for internal grants tend to be identical (for the CS department for me, they are all named "David Doty ENGR COMPUTER SCIENCE PPM Only").

To clean up the project names, you can specify two arguments to the `Summary.from_file` method: `substrings_to_clean` and `suffixes_to_clean`. Any substring appearing in `substrings_to_clean` will be removed, for example if I set `substrings_to_clean=['CS', 'Doty']`, it will change the project name `"CS NSF Engineering DNA and RNA Doty K302325F33"` to `"NSF Engineering DNA and RNA K302325F33"`. Anything in `suffixes_to_clean` will be removed, not only that substring, but the entire rest of the name. For instance, if I set `suffixes_to_clean=['K3023']`, it will change `"NSF Engineering DNA and RNA  K302325F33"` to `"NSF Engineering DNA and RNA"`.

I personally use them like this:
```python
suffixes = ['K3023', 'DOTY DEFAULT PROJECT 13U00']
substrings = ['Doty', 'CS ']
summary_aug = Summary.from_file('2024-8-1.xlsx', substrings_to_clean=substrings, suffixes_to_clean=suffixes)
```
due to the particular manner in which someone mashed their palm against the keyboard to generate the alien-looking project names of my own grants (e.g. "*CS NSF DNA and RNA Partic Support Doty K3023EDRNA*"), but you will want to customize according to the shape of your SPO representative's palm.

## API
[API documentation](https://aggie-unterprise.readthedocs.io/).


## Standalone program
If you do not want to install Python or write Python code, there is a standalone command-line program called aggie-report that can do some basic tasks. There are pre-compiled versions of aggie-report you can download:
- [Windows](https://github.com/dave-doty/aggie-unterprise/releases/latest/download/aggie-report-win.exe) (save with file name `aggie-report.exe`)
- [Linux](https://github.com/dave-doty/aggie-unterprise/releases/latest/download/aggie-report-linux) (save with file name `aggie-report`)
- [macOS](https://github.com/dave-doty/aggie-unterprise/releases/latest/download/aggie-report-mac) (save with file name `aggie-report`)

Open a command prompt in the directory where you saved the file.

On Mac (respectively Linux), type `mv aggie-report-mac aggie-report` to rename the file to `aggie-report` (respectively `mv aggie-report-linux aggie-report` on Linux).

On Mac and Linux, type `chmod +x aggie-report` (after renaming the file to `aggie-report`) to make the program executable.

On Mac, you must then type `xattr -d com.apple.quarantine ./aggie-report` to enable the program to run.

Run `aggie-report -h` to see all the options.

Alternately, you can [install Python](https://www.python.org/downloads/), then install the package by typing `pip install aggie_unterprise`. After doing this, you will have a program named `aggie-report` you can use from any directory:

```
$ aggie-report -d reports
```

Assuming you have a subdirectory `reports` of your current directory with your AggieEnterprise reports in it, this will print to the screen a summary of each report, as well as a summary of differences between adjacent-in-creation-time reports. It sorts them in descending order of the date they were produced (so prints the latest one first). 

The other options are as follows, but in case this documentation gets outdated, run `aggie-report -h` to see the latest options:

```
usage: aggie-report [-h] [-d DIRECTORY | -f FILE [FILE ...]] [-o OUTFILE] [-nd | -ni] 
[-s] [-sb SUBSTRING [SUBSTRING ...]] [-sf SUFFIX [SUFFIX ...]] 
[-sbf SUBSTRINGS_FILE] [-sff SUFFIXES_FILE]

Processes reports generated by AggieEnterprise to summarize the useful data
in them. By default it sorts the files by the date they were generated
(according to cell A3 inside the spreadsheet file), and in that order
going backwards (so latest file is processed first), summarizes each file,
as well as summarizing differences between adjacent files.
If run with no arguments like this:

    C:\reports> python aggie_report.py

or if you are using the executable aggie-report included with the
aggie_unterprise Python package, like this:

    C:\reports> aggie-report

it will process all the .xlsx files in the current directory and print the
results to the screen. Command line arguments can be used to customize this;
type  `python aggie_report -h` to see all the options.

If a file is in the current directory named `names_to_clean.json`, in a format
like this:

    {
        "Long ugly project name like NSF CAREER K20304932": "CAREER",
        "Another long ugly project name like NSF Small K302777": "Small"
    }

then the program will clean up the project names by replacing the projects
names on the left side of the `:` with the cleaner project name on the right
side of the `:`. This is more flexible than the -sb and -sf options.
If a project name appears in the JSON file, then it is not processed at all
using the -sb and -sf options. (Nor the file-based options --sbf and --ssf.)

options:
  -h, --help            show this help message and exit
  -d DIRECTORY, --directory DIRECTORY
                        directory to search for .xlsx files generated by AggieEnterprise. 
                        All .xlsx files are processed; to process only some files, 
                        list them explicitly after the flag -f. If neither -d nor -f 
                        is given, all .xlsx files in the current directory are processed. 
                        This option is mutually exclusive with -f.
  -f FILE [FILE ...], --files FILE [FILE ...]
                        List of one or more .xlsx files generated by AggieEnterprise. 
                        To specify that all .xlsx files in a directory should be 
                        processed, use the -d option. This option is mutually exclusive
                        with -d.
  -o OUTFILE, --outfile OUTFILE
                        Name of file to print output; if not specified, print output to 
                        the screen.
  -nd, --no-diffs       If specified, do not include differences between adjacent files 
                        in the output. This option is mutually exclusive with -ni.
  -ni, --no-individual  If specified, do not include summaries for individual files. 
                        This option is mutually exclusive with -nd.
  -s, --sort-increasing-by-date
                        If specified, sort the files by date in increasing order instead 
                        of the default, which is to sort by date in decreasing order.
  --sb SUBSTRING [SUBSTRING ...], --substrings SUBSTRING [SUBSTRING ...]
                        List of substrings to remove from the project name. This can
                        help clean up ugly project names like "NSF CAREER K20304932"
                        by specifying substring "K20304932", which would change the
                        project name to "NSF CAREER".
  --sf SUFFIX [SUFFIX ...], --suffixes SUFFIX [SUFFIX ...]
                        List of substrings to remove from the project name, as well
                        as the entire suffix from there to the end. This can help
                        clean up ugly project names like "NSF CAREER K302777" and
                        "NSF Small K302999" by specifying substring "K302", which
                        would change these project names "NSF CAREER" and "NSF
                        Small".
  --sbf SUBSTRINGS_FILE, --substrings-file SUBSTRINGS_FILE
                        Filename containing substrings (separated by whitespace or
                        newlines) to remove from the project name. This is like the
                        -sb option, but reads the substrings from a file so that you
                        do not have to type them all out at the command line and can
                        reuse them in multiple runs.
  --ssf SUFFIXES_FILE, --suffixes-file SUFFIXES_FILE
                        Filename containing substrings (separated by whitespace or
                        newlines) to remove from the project name, as well as the
                        entire suffix from there to the end. This is like the -sf
                        option, but reads the substrings from a file so that you do
                        not have to type them all out at the command line and can
                        reuse them in multiple runs.
  -c, --show-cents      If specified, show dollar amounts including cents; default 
                        behavior is to round to the nearest dollar.
```

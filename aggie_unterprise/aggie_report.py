from dataclasses import dataclass, field
import argparse
import sys
import os
import json
import itertools
from pathlib import Path
from typing import Optional, List, TextIO
from aggie_unterprise import Summary


def main():
    args: CLArgs = parse_command_line_arguments()
    paths = find_filenames(args)
    print('Summarizing grant data from AggieEnterprise reports in these files:')
    print('  ' + ', '.join(path.name for path in paths))
    print(f'Output will be written to the ' + (f'file {args.outfile}' if args.outfile is not None else 'screen'))

    names_to_clean = read_names_json('names_to_clean.json')
    substrings = combine_substrings_from_file(args.substrings_file, args.substrings_to_clean)
    suffixes = combine_substrings_from_file(args.suffixes_file, args.suffixes_to_clean)

    summaries = [Summary.from_file(path, names_to_clean=names_to_clean,
                                   substrings_to_clean=substrings, suffixes_to_clean=suffixes) for path in paths]
    summaries.sort(key=lambda s: s.date_and_time)
    if not args.sort_increasing_by_date:
        summaries.reverse()

    # For some reason using a context manager causes errors when writing to stdout.
    # Normally it works, but when I use PyInstaller to create an executable, it tries to close stdout
    # and causes an error
    file = open(args.outfile, 'w', encoding='utf-8') if args.outfile is not None else sys.stdout
    try:
        for sum_prev, sum_next in itertools.pairwise(summaries):
            if not args.sort_increasing_by_date:
                sum_prev, sum_next = sum_next, sum_prev
            if args.include_individual_summaries:
                print_indv_summary(file, sum_prev if args.sort_increasing_by_date else sum_next, args.show_cents)
            if args.include_diffs:
                print_diff_summary(file, sum_prev, sum_next, args.show_cents)
        if args.include_individual_summaries:
            last_summary = summaries[-1]
            print_indv_summary(file, last_summary, args.show_cents)
    finally:
        if file is not sys.stdout:
            file.close()

def combine_substrings_from_file(substrings_file: Optional[str], substrings: List[str]) -> List[str]:
    if substrings_file is not None:
        substrings = list(substrings)
        with open(substrings_file, 'r', encoding='utf-8') as file:
            substrings.extend(file.read().split())
    return substrings


def read_names_json(fn: str) -> dict[str, str]:
    '''
    Check if 'names_to_clean.json' exists, read it into a Dict[str, str], and validate format.

    Returns:
        Dictionary of name mappings if file exists and has valid format, empty dict otherwise.
    '''
    if not os.path.isfile(fn):
        return {}

    with open(fn, 'r') as file:
        # Parse JSON into a Python dictionary
        data = json.load(file)

        # Validate that data is a dictionary
        if not isinstance(data, dict):
            print("Error: JSON content is not a dictionary")
            return {}

        # Validate that all values are strings
        for key, value in data.items():
            if not isinstance(value, str):
                print(f'Error: Value "{value}" for key "{key}" is not a string but is type {type(value)}')
                return {}

        # Confirm successful loading
        print(f"Successfully loaded dictionary with {len(data)} entries")
        return data


def print_diff_summary(file: TextIO, sum_prev: Summary, sum_next: Summary, show_cents: bool) -> None:
    print(f'\nDifferences from {sum_prev.date()} to {sum_next.date()}:', file=file)
    print(sum_next.diff_table(sum_prev, show_cents=show_cents), file=file)


def print_indv_summary(file: TextIO, sum_prev: Summary, show_cents: bool) -> None:
    print(f'\nTotals for {sum_prev.date()}:', file=file)
    print(sum_prev.table(show_cents=show_cents), file=file)


@dataclass
class CLArgs:
    directory: Optional[str] = None
    # Directory containing .xlsx files generated by aggie

    files: Optional[List[str]] = None
    # Files specified by user in lieu of specifying directory

    outfile: Optional[str] = None
    # Name of file to print output; if not specified, output to the screen

    include_diffs: bool = True
    # Whether to include differences between files in the output

    include_individual_summaries: bool = True
    # Whether to include individual summaries in the output

    sort_increasing_by_date: bool = False
    # Whether to sort the files by date in increasing order

    substrings_to_clean: List[str] = field(default_factory=list)
    # List of substrings to remove from the project name

    suffixes_to_clean: List[str] = field(default_factory=list)
    # List of substrings to remove from the project name, including
    # the whole suffix following the substring

    substrings_file: Optional[str] = None
    # Filename containing substrings to remove from the project name

    suffixes_file: Optional[str] = None
    # Filename containing substrings to remove from the project name,
    # including the whole suffix following the substring

    show_cents: bool = False
    # Whether to show cents in the output. If False, round to the nearest dollar.


def parse_command_line_arguments() -> CLArgs:
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description=r'''Processes reports generated by AggieEnterprise to summarize the useful data 
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
''',
    )
    # formatter_class=argparse.RawTextHelpFormatter)

    dir_files_group = parser.add_mutually_exclusive_group(required=False)

    dir_files_group.add_argument('-d', '--directory', type=str,
                                 help='''\
directory to search for .xlsx files generated by 
AggieEnterprise. All .xlsx files are processed; to 
process only some files, list them explicitly after 
the flag -f. If neither -d nor -f is given, all 
.xlsx files in the current directory are processed.
This option is mutually exclusive with -f.''')

    dir_files_group.add_argument('-f', '--files', nargs='+', type=str, metavar='FILE',
                                 help='''\
List of one or more .xlsx files generated by 
AggieEnterprise. To specify that all .xlsx files in 
a directory should be processed, use the -d option. 
This option is mutually exclusive with -d.''')

    parser.add_argument('-o', '--outfile', type=str,
                        help='''\
Name of file to print output; if not specified, 
print output to the screen.''')

    no_summary_group = parser.add_mutually_exclusive_group(required=False)

    no_summary_group.add_argument('-nd', '--no-diffs', action='store_true',
                                  help='''\
If specified, do not include differences between 
adjacent files in the output. This option is
mutually exclusive with -ni.''')

    no_summary_group.add_argument('-ni', '--no-individual', action='store_true',
                                  help='''\
If specified, do not include summaries for 
individual files. This option is mutually
exclusive with -nd.''')

    parser.add_argument('-s', '--sort-increasing-by-date', action='store_true',
                        help='''\
If specified, sort the files by date in increasing
order instead of the default, which is to sort by
date in decreasing order.''')

    parser.add_argument('--sb', '--substrings', dest='substrings', nargs='+', type=str, metavar='SUBSTRING',
                        help='''\
List of substrings to remove from the project name.
This can help clean up ugly project names like 
"NSF CAREER K20304932" by specifying substring "K20304932",
which would change the project name to "NSF CAREER".''')

    parser.add_argument('--sf', '--suffixes', dest='suffixes', nargs='+', type=str, metavar='SUFFIX',
                        help='''\
List of substrings to remove from the project name,
as well as the entire suffix from there to the end.
This can help clean up ugly project names like 
"NSF CAREER K302777" and "NSF Small K302999" 
by specifying substring "K302", which would change 
these project names "NSF CAREER" and "NSF Small".''')

    parser.add_argument('--sbf', '--substrings-file', dest='substrings_file', type=str,
                        help='''\
Filename containing substrings (separated by whitespace 
or newlines) to remove from the project name. This is 
like the -sb option, but reads the substrings from a file
so that you do not have to type them all out at the 
command line and can reuse them in multiple runs.''')

    parser.add_argument('--ssf', '--suffixes-file', dest='suffixes_file', type=str,
                        help='''\
Filename containing substrings (separated by whitespace 
or newlines) to remove from the project name, as well as
the entire suffix from there to the end. This is 
like the -sf option, but reads the substrings from a file
so that you do not have to type them all out at the 
command line and can reuse them in multiple runs.''')

    parser.add_argument('-c', '--show-cents', action='store_true',
                                  help='''\
If specified, show dollar amounts including cents; 
default behavior is to round to the nearest dollar.''')

    args = parser.parse_args()
    clargs = CLArgs()

    if args.directory is not None:
        assert not args.files
        clargs.directory = args.directory

    if args.files is not None:
        assert not args.directory
        clargs.files = args.files

    if args.directory is None and args.files is None:
        clargs.directory = '.'

    if args.outfile:
        clargs.outfile = args.outfile

    clargs.include_diffs = not args.no_diffs
    clargs.include_individual_summaries = not args.no_individual
    assert clargs.include_diffs or clargs.include_individual_summaries

    clargs.sort_increasing_by_date = args.sort_increasing_by_date

    if args.substrings is not None:
        clargs.substrings_to_clean = args.substrings

    if args.suffixes is not None:
        clargs.suffixes_to_clean = args.suffixes

    if args.substrings_file is not None:
        clargs.substrings_file = args.substrings_file

    if args.suffixes_file is not None:
        clargs.suffixes_file = args.suffixes_file

    clargs.show_cents = args.show_cents

    return clargs


def find_filenames(args) -> List[Path]:
    if args.files is not None:
        return [Path(file).resolve() for file in args.files]
    assert args.directory is not None
    directory = Path(args.directory).resolve()
    fns = [fn for fn in directory.iterdir() if fn.suffix == '.xlsx']
    if len(fns) == 0:
        raise FileNotFoundError(f'No .xlsx files found in directory "{directory}"')
    return fns


if __name__ == '__main__':
    main()

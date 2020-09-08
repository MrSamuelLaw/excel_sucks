#!/usr/bin/env python3
import re
import argparse
from pathlib import PurePath


"""allows the user write an excel equation in a .txt file
then convert that into a single line equation which can be pasted in the
excel function bar.
Also allows the user to convert relative references into
"indirect" function. see the relative_to_indirect function for
more info."""


def cli():
    """walks the user through creating a function from
    a file assuming they have the path to the file"""

    filepath = input("copy and paste file path, then press enter\n")
    eqn = format_from_file(str(filepath))
    option1 = input('would you like to make the function "indirect"? [Y,N]\n')
    if option1.upper() == "Y":
        option2 = input('does this function need to work on every row or column? [row, col]\n')
        if option2.lower() == "row":
            print(relative_to_indirect(eqn, row_or_col="row"))
        elif option2.lower() == "col":
            print(relative_to_indirect(eqn, row_or_col="col"))
        else:
            raise ValueError(f'options are row or "col", "not" {option2}')
    else:
        print(eqn)


def format_from_file(filepath: str) -> str:
    """Asks the user to drag and drop file
    into terminal and returns the flattened text"""

    path = PurePath(filepath)
    with open(path, "r") as f:
        return flatten(str(f.read()))


def flatten(text: str) -> str:
    """Takes all of the Excel logic spread accross
    multiple lines for readability and returns a single
    line for pasting into the excel function bar"""

    results = []
    for line in text.splitlines():
        line = line.lstrip()
        line = line.rstrip()
        results.append(line)
    return "".join(results)


def relative_to_indirect(equation: str, row_or_col: str = "row") -> str:
    """One of the major problems with web excel is
    that the cells cannot be locked, and when using relative
    references, it only takes one accidental copy and paste
    to ruin, an entire workbook, to prevent this, this function
    will take a function with relative references and convert them
    to indirect. For example:
    LEN(A5), you want to copy and paste this formula to all the rows
    and have it work for A(any row here) then you would have to write
    INDIRECT(CONCAT("A", ROW())) which is time consuming and error prone
    This function saves you the hassle by converting every indirect
    cell refernece to the INDIRECT format to then be copy and pasted
    into the cells of a web excel document

    args:
        equation: equation to have relative reference converted to
            indirect.
        row_or_col: makes either the row or column reference dynamic.
            i.e. row = "A", ROW() or col = "COL()", 5.
            valid inputs are "row" or "col"
            defaults to row
    """
    # define pattern
    pattern = re.compile(r"[A-Z]{1,3}\d{1,5}")
    p = re.compile(r"[A-Z]")

    # define splitting function
    def split(match: str):
        idx, span = p.search(match).span()
        return match[:span], match[span:]

    if row_or_col == "row":
        matches = pattern.findall(equation)
        for m in matches:
            static, _ = split(m)
            equation = equation.replace(
                    m,
                    f'INDIRECT(CONCAT("{static}", ROW()))'
                )
    elif row_or_col == "col":
        matches = pattern.findall(equation)
        for m in matches:
            _, static = split(m)
            equation = equation.replace(
                    m,
                    f'INDIRECT(CONCAT(COLUMN(), "{static}"))'
                )
    else:
        raise ValueError(
            f"""row_or_col value is invalid, options
             are "row" or "col", not {row_or_col}"""
        )

    return equation


if __name__ == "__main__":
    """runs the script"""
    cli()

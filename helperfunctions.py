import pandas as pd
from pathlib import Path

def uniq(lst):
    last = object()
    for item in lst:
        if item == last:
            continue
        yield item
        last = item


def sort_and_deduplicate(l):
    return list(uniq(sorted(l, reverse=True)))


def remove_null(list):
    for i in list:
        if i == "nan" or i == '' or i is None or pd.isna(i) == True:
            list.remove(i)
    return list


def sheet_sort_rows(ws, row_start, row_end=0, cols=None, sorter=None, reverse=False):
    """ Sorts given rows of the sheet
        row_start   First row to be sorted
        row_end     Last row to be sorted (default last row)
        cols        Columns to be considered in sort
        sorter      Function that accepts a tuple of values and
                    returns a sortable key
        reverse     Reverse the sort order
    """

    bottom = ws.max_row
    if row_end == 0:
        row_end = ws.max_row
    right = get_column_letter(ws.max_column)
    if cols is None:
        cols = range(1, ws.max_column + 1)

    array = {}
    for row in range(row_start, row_end + 1):
        key = []
        for col in cols:
            key.append(ws.cell(row, col).value)
        array[key] = array.get(key, set()).union({row})

    order = sorted(array, key=sorter, reverse=reverse)

    ws.move_range(f"A{row_start}:{right}{row_end}", bottom)
    dest = row_start
    for src_key in order:
        for row in array[src_key]:
            src = row + bottom
            dist = dest - src
            ws.move_range(f"A{src}:{right}{src}", dist)
            dest += 1

def check_and_create_path( path: Path):
    path_way = path.parent if path.is_file() else path

    path_way.mkdir(parents=True, exist_ok=True)

    if not path.exists():
        path.touch()


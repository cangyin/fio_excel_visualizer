
# import logging
# logging.basicConfig(level=logging.DEBUG, format='%(asctime)s [%(module)s:%(lineno)d][%(levelname)s] %(message)s')
# logger = logging.getLogger('main')

import os
import re
import json
import csv
import glob
from itertools import accumulate
from types import SimpleNamespace

from excel_utils import ExcelUtils, ExcelXYChartUtils, app


## CSV Utilities

def read_csv(path :str, dialect='excel'):
    # csv.excel
    csv_reader = csv.reader(open(path, 'r'), dialect=dialect)
    return [ [ x.strip() for x in row ] for row in csv_reader ]


def save_as_csv(table, path, dialect='excel'):
    with open(path, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile, dialect=dialect)
        for row in table:
            csv_writer.writerow(row)


## Miscellaneous Utilities

def glob_a_file(pattern) -> str:
    files = glob.glob(pattern)
    if len(files) != 1:
        raise f'number of files globed is not 1: log_file={repr(files)}'
    return files[0]


def read_json_as_object(json_file):
    def json_object_hook(d :dict):
        # replace ' ' in keys with '_'
        return SimpleNamespace(**{ k.replace(' ', '_'): v for k, v in d.items() })
    with open(json_file, 'r') as f:
        return json.load(f, object_hook=json_object_hook)


## Series Manipulation Utilities

# Define 'series':
#   A 2-dimensional array. Each element of 'series' is a 'row', each 'row' is a
#   list of primary python values. Length of each 'row' is the same. For example,
#   [ [1, 2], [3, 4], [5, 6], ... ]

def read_series(file, filter=None, columns=None, headers=None, no_empty=True):
    if (columns and headers) and (len(columns) != len(headers)):
        raise Exception('len(columns) != len(headers)')

    series = read_csv(file)
    if filter:
        series = [ row for row in series if filter(row) ]
        if len(series) == 0:
            return [headers] if no_empty else []
    if columns:
        series = [ [ row[i] for i in columns ] for row in series ]
    if headers:
        prepend_header(series, headers)
    return series


get_bw_series   = lambda file, ddir=0: read_series(file, filter=lambda row: row[2]==str(ddir), columns=[0, 1, 2], headers=['Time', 'Bandwidth', 'Cumulative'])
get_iops_series = lambda file, ddir=0: read_series(file, filter=lambda row: row[2]==str(ddir), columns=[0, 1], headers=['Time', 'IOPS'])
get_lat_series  = lambda file, ddir=0: read_series(file, filter=lambda row: row[2]==str(ddir), columns=[0, 1], headers=['Time', 'Latency'])


def prepend_header(series, header):
    if isinstance(header, list):
        header.extend([''] * (len(series[0]) - len(header)))
    elif isinstance(header, str):
        header = [header] + [''] * (len(series[0]) - 1)
    series.insert(0, header)
    return series


def extend_series(series, axis=0, to=0, fill_value=''):
    if axis == 0: # horizontal
        if to > len(series[0]):
            padding = [ fill_value ] * (to - len(series))
            series = [ row + padding for row in series ]
    elif axis == 1: # vertical
        if to > len(series):
            series.extend([ [fill_value] * len(series[0]) ] * (to - len(series)))
    else:
        raise Exception('axis must be 0 or 1')
    return series


def concatenate(*list_of_series, delimited=True):
    # filter out empty series
    list_of_series = [ series for series in list_of_series if series ]

    # align all series
    max_len = max([ len(s) for s in list_of_series ])
    for series in list_of_series:
        series.extend([ [''] * len(series[0]) ] * (max_len - len(series)))
    
    # concatenate
    if delimited:
        return [ [ c for sublist in row for c in sublist + [''] ][0:-1] for row in zip(*list_of_series) ]
    else:
        return [ [ c for sublist in row for c in sublist ] for row in zip(*list_of_series) ]


## The major functions

def aggregate(o :SimpleNamespace) -> dict:
    '''
    According to the json (output of Fio) object, read data from log files, and return a
    list of tables. Each table corresponds to a job's log data, and consists of 3 groups
    of series, where series are grouped by data direction (read, write, trim). Series in a
    group are simply concatenated.

    A table is like:
    __________________________________________________________________________________________
    |                 Read                        |        Write        |        Trim        |
    | Time | BW  | Cum | Time | IOPS | Time | Lat |         ...         |        ...         |
    |      |     |     |      |      |      |     |         ...         |        ...         |
    |      |     |     |      |      |      |     |         ...         |        ...         |
    ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

    All log files should be found in current directory.
    '''
    tables = {
        # job_name: table_data
    }

    _tag_lambdas = {
        'bw': get_bw_series,
        'iops': get_iops_series,
        'lat': get_lat_series,
    }

    for job in o.jobs:
        rw = getattr(job.job_options, 'rw') or getattr(job.job_options, 'readwrite') or None
        if rw is None:
            print(f'skipped job {job.jobname} for it has no rw type')
            continue

        # all series for this job, grouped to be a table.
        table = { }

        for ddir in (0, 1, 2): # read, write, trim
            # skip if no I/O operations in this direction
            _keywords = (('read', 'rw'), ('write', 'rw'), ('trim', ))[ddir]
            if not any([ s in rw for s in _keywords ]):
                continue

            data_direction = ('Read', 'Write', 'Trim')[ddir]

            ddir_group = [] # all series for this data direction
            for tag, _get_series  in _tag_lambdas.items():
                log_file = glob_a_file(f'{job.jobname}_{tag}.*.log')

                s = _get_series(log_file, ddir)
                s.insert(1, ['0'] * len(s[0])) # insert all '0' row
                # concatenate 2 series to a new series
                ddir_group = concatenate(ddir_group, s, delimited=True)

            table[data_direction] = prepend_header(ddir_group, data_direction)

        ## additional series goes here.
        if table:
            # `data_direction` should have been defined, if `table` is not empty.
            ddir_group = table[data_direction]
            # table[data_direction] = concatenate(ddir_group, <additional series>)

        # series = prepend_header(series, f"Job '{job.jobname}'")
        tables[job.jobname] = table
        save_as_csv(table, f'{job.jobname}.csv')

    return tables


def aggregate_and_visualize(json_file :str, util :ExcelUtils, sheet_name=None, table_offset=0):
    '''
    sheet_name:
        if given, tables are stored in a single sheet, else each table uses one sheet.
    table_offset:
         column offset in sheet, from where to start filling data.
    '''
    o = read_json_as_object(json_file)

    # change dir to the directory of json and log files.
    old_cwd, new_cwd = os.getcwd(), os.path.split(json_file)[0] or '.'
    os.chdir(new_cwd)

    tables = aggregate(o)
    del o

    for jobname, table in tables.items():

        ddir_group_widths = [ len(g[0]) for g in table.values() ]
        # accumulate([1,2,3,4,5], initial=100) --> 100 101 103 106 110 115 (one more element than input)
        ddir_group_offsets = list(accumulate(ddir_group_widths, initial=table_offset))

        ## let's move on to Excel steps.
    
        # get a worksheet
        sheet = util.get_sheet(sheet_name or jobname)
        sheet.Activate() # make it the ActiveSheet
        
        # dump table to excel sheet.
        for offset, ddir_group in zip(ddir_group_offsets, table.values()):
            # replace empty str with None
            ddir_group = tuple( tuple(x or None for x in row) for row in ddir_group )
            util.fill_cells(1, offset + 1, ddir_group)

        # 自适应列宽
        app.Columns.AutoFit()

        # formulas: cumulative_IO_size
        # = (current_time - previous_time) * current_BW + cumulated_IO_size
        formula = "=(((RC[-2]-R[-1]C[-2])/1000)*RC[-1]/1024/1024+R[-1]C)"
        for offset, ddir_group in zip(ddir_group_offsets, table.values()):
            formula_col = offset + 3
            rbox = [(4, formula_col), (len(ddir_group), formula_col)]
            util.set_formula( rbox, formula )
            util.range(rbox).NumberFormatLocal = "0.00_ " # two decimal places

        # charts
        for offset, data_direction in zip(ddir_group_offsets, table.keys()):

            chart = ExcelXYChartUtils.create_with_declaration(sheet, {
                "title": f"{data_direction} Performance in Job '{jobname}'",
                "series":  [
                    [ (4, offset + 1, offset + 2), 1,  "Bandwidth" ],
                    [ (4, offset + 8, offset + 9), 2,  "Latency" ],
                ],
                "x_title": "msec",
                "y_titles": [ "KB/s", "nsec" ],
                "position": (util.range([(1, 1), (1, offset + 1)]).Width, util.range([(1, 1), (2, 1)]).Height),
                "with_chart": f".Parent.Width = 450",
            }).chart

            chart = ExcelXYChartUtils.create_with_declaration(sheet, {
                "title": f"{data_direction} Performance by Cumulative I/O Size in Job '{jobname}'",
                "series":  [
                    [ (4, offset + 3, offset + 2), 1,  "Bandwidth" ],
                ],
                "x_title": "GB",
                "y_titles": [ "KB/s", "" ],
                "position": (util.range([(1, 1), (1, offset + 1)]).Width, chart.Parent.Top + chart.Parent.Height + 1),
                "with_chart": f".Parent.Width = 450",
            }).chart

        # TODO: histograms: latency

        # TODO: common statistics: avg of bw, avg of iops, avg of latency, min max values.

        if sheet_name:
            table_offset += ddir_group_offsets[-1] + 1 # 隔一列

    os.chdir(old_cwd)
    return None


if __name__ == '__main__':
    workbook = app.Workbooks.Add()

    util = ExcelUtils(workbook)

    json_file = 'tmp2/jfs-fio.json'

    aggregate_and_visualize(json_file, util, sheet_name="tmp2")

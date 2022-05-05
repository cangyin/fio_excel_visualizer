# Requirements
- pywin32


# Command Line Help

```sh
usage: fio_excel_visualizer.py [-h] [--sheet-name SHEET_NAME]
                               [--sheet-names [SHEET_NAMES [SHEET_NAMES ...]]] [-o [O]] [--close]
                               dirs [dirs ...]

Aggregate and visualize log data of Fio output.

positional arguments:
  dirs                  Directories of json and log files output by Fio. Each directory contains all log     
                        and json output generated during a Fio workload.

optional arguments:
  -h, --help            show this help message and exit
  --sheet-name SHEET_NAME
                        Name of the sheet to store the data. If not given, data in each directories will be  
                        stored in separate sheets, where sheet name is the directory name.
  --sheet-names [SHEET_NAMES [SHEET_NAMES ...]]
                        Name of the sheets to store the data. Number of the names should be the same as      
                        that of `dirs`.
  -o [O]                save the file with the given pathname.
  --close               Close the file after saving. Must be used with `-o`.
```

# General Usage Steps

1. Run Fio workload, with [output format](https://fio.readthedocs.io/en/latest/fio_man.html#cmdoption-output-format) set to JSON/JSON+ and enable [log outputs](https://fio.readthedocs.io/en/latest/fio_man.html#measurements-and-reporting) (typically `write_bw_log`, `write_lat_log`, `write_iops_log`, `log_avg_msec`).
2. Assume the json and log files generated by Fio are located in directory named `stats`.
3.  Run the script from parent directory of `stats`: `python fio_excel_visualizer.py stats/`


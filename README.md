# Database_Architecturing_Documentation_Tool

This *validator tool* supports metadata auditing efforts by automating the validation of documented metadata against the corresponding data available in Dremio. Its primary purpose is to identify inconsistencies between the documentation and the corresponding Dremio records, reducing the effort required for manual validation. These *inconsistencies/errors* are highlighted in the file's output to display exactly where each discrepancy is, and a specified color that corresponds to the type of flag being raised. The tool primarily operates on the *Full Mapping* worksheet, although its configuration-driven design allows for adaptability to evolving business requirements and alternative Excel workbook formats.

## Error Classification Matrix

| Error Code Group | Flag Type | Color |
| ---------------- | --------- | ----- |
| `100` errors | Blank fields, exception of `N/A` values | 🔴 Red |
| `200` errors | Source-to-Target mismatch | 🟣 Purple |
| `300` errors | Target-to-Dremio mismatch | 🔵 Blue |
| -c f | Colorblind argument that mutes colors | ⚫ Grey |

<br><br>

# Quick Start

```bash
# All Args Ex.
python Main.py --filepath <Files> --output <Folder> --user <UsernameID> --env <uat|prod> --test <1|2|3> --colorblind <t|f> --details <c|d|f>
python Main.py -f <Files> -o <Folder> -u <UsernameID> -e <uat|prod> -t <1|2|3> -c <t|f> -d <c|d|f>
```

The only required arguments is `f`/`--filepath`, the path to the file(s) that will be analyzed in this script.

## Folder(s)/File Args

```bash
# -f, --filepath
python Main.py --filepath <Files>
python Main.py -f <Files>

# -o, --output
python Main.py --filepath <Files> --output <Folder>
python Main.py -f <Files> -o <Folder>

# -t, --test
python Main.py --filepath <Files> --test <1|2|3>
python Main.py -f <Files> -t <1|2|3>
```

Additionally, users who prefer to specify where to receive all validated files can use the `-o`/`--output`. Files can be received in the same folder as the argument for ingestion, `-f`/`--filepath`. All output files are marked with *RESULTS_X_* prefix to distinguish from older versions. The *X* is a counter that continuously increasing every time the output already has the prefix *RESULTS_*. To perform a test, users must specify the `-t`/`--test` argument, with any combination of the three available options `1|2|3`. On the contrary, users do not have to run any test, in this case the script will only utilize the `formatter()` function which is configurable. __Important Note:__ The order in which error codes are executed is dependent on user specifiication on the command line. This matters in the case that one particular cell is flagged with several different errors. *Recommended* to always layer error 1 last as provided in the example.

## Dremio Args

```bash
# -u, --user
python Main.py --filepath <directory|file> --user <username>
python Main.py -f <directory|file> -u <username>

# -e, --env
python Main.py -filepath <directory|file> --env <uat|prod>
python Main.py -f <directory|file> -e <uat|prod>
```

Arguments pertaining to Dremio access is the username and the environment selection, however these are ooptional in the command line. This is because the only test that requires a connection to dremio is `300` *error*. Otherwise, the script will not establish a connection to dremio allowing for more efficient run-time. If the username argument is not specified, the terminal will prompt users in th eterminal to enter all their credentials. For repetitive runs, it is recommended to utilize the `-u`/`--user` arg, as to prevent typos. In terms of environment, the default is always *UAT*, users must specifiy for the script to run in *PROD*.

## Other Args

```bash
# -c, --colorblind
python Main.py -filepath <directory|file> --colorblind <t|f>
python Main.py -f <directory|file> -c <t|f>

# -d, --details
python Main.py -filepath <directory|file> --details <c|d|f>
python Main.py -f <directory|file> -d <c|d|f>
```

Users who are color-blind or color-sensitive can use the `-c` / `--colorblind` argument to disable colored output. Available options are `t=terminal output` and `f=file output`. Users may specify either option individually or combine them to disable coloring in both locations. Commas are optional when combining options (for example, `tf` and `t,f` are treated the same). When enabled, all colorized output is removed except for workbook schema formatting. __Important Note:__ When file coloring is disabled (`f`), validation flags in generated files are displayed in gray. For the best visibility, it is recommended to run one error-reporting option at a time.

The validator reports metadata ingestion and Dremio view ingestion issues through terminal notifications. Users can use the `-d` / `--details` argument to suppress specific categories of output when a large number of notifications are generated. To suppress a category, users usimply specify th eoptions that correlate to which notifications they want to receive. Available options that correlate to which notifications they want to receive. Available options are `c=column checker`, `d=Dremio views`, and `f=failed files summary`. Multiple options may be dcombined, with or without commas, depending on the desired level of output. The `f` option controls the terminal summary of failed files. In most cases, it is recommended to either suppress both `c` and `d` notifications together or suppress only `f` to reduce terminal noise while maintaining useful validation feedback.

# Requirements/File(s) Handling

### Failed Files

The ingestion of all files are handled independently. Failure to process a particular file will not stop the script from executing, as it will continue for all remaining files. At the end of execution, if any files failed to be processed through the script, a *Failed Files* summary will be reported in the terminal. The only exception is Dremio connection validation. If the *Dremio connection test fails, execution is terminated* before data ingestion begins.

### Metadata

- Excel workbook (`.xlsx` | `.xls`).
- *Full Mapping* worksheet, otherwise refer to [Worksheet Selection](#worksheet-selection-map_sheet).
- Worksheet schema alignment with configuration file, otherwise refer to [Schema Configuration](README_Config.md/#Excel-Workbook-Schema-Configuration).
- Proper Dremio credentials & entitlements, only required when executing `error 3`.

### Validation Outputs

- *Metadata* worksheet contains the following:
  - [Color-coded validation flag](README_Config.md/#Error-Classification-Matrix).
  - Fully-formatted table based on [configuration file](README_Config.md/#Excel-Workbook-Schema-Configuration).
  - All other sheets included from original input file.
- *Extra Errors* worksheet contains the following:
  - The validation-[color legend](README_Config.md/#Excel-Workbook-Schema-Configuration).
  - Additional context, such as missing views, missing fields, or DremioSQL file paths.
- *RESULTS_X_*
  - All files will have a naming convention that contains an incremental counter.
  - Allows for same workbook to be modified and revalidated.

<br><br>

# Script Workflow

| File | Responsibilities |
| ---- | ---------------- |
| `Main.py` | Orchestrates the end-to-end workflow and coordinates interactions between all modules |
| `DremioConnection.py` | Serves as the driver which connects the python scripting to the Dremio Data Platform, handling data retrieval |
| `ProcessMetadata.py` | Manages workbook ingestion, standardization, validation execution, and result preparation |
| `ErrorCodes.py` | Defines business validation results into a readable, formatted workbook |
| `Utilities.py` | Centralizes utility functions and application constants |
| `Configs.toml` | All configuration details (workbook shcema, error codes, conversions) |
| `requirements.txt` | *pandas*, *openpyxl*, *pyarrow*, *pathlib*, *tqdm*, etc. |

## File Dependencies

```text
                    Main.py -------------------------------+
                       |                                   |
          +------------+------------+                      |
          |                         |                      |
Dremio Connection.py       ProcessMetadata.py              |
                                    |                      |
                       +------------+------------+         |
                       |            |            |         |
                  Formatter.py      |      ErrorCodes.py   |
                       |            |            |         |
                       +------------+------------+         |
                                    |                      |
                               Utilities.py ---------------+
                                    |
                               Configs.toml
```

## Execution Sequence

1. Verifying all given arguments.
2. If running `--test <3>`, test the Dremio Connection and credentials. 
3. Ingest metadata workbook(s).
4. Run initial validation of the worksheet structure and standardize metadata.
5. If running `--test <3>`, read and ingest Dremio views from the workbook's given paths.
6. If any `--test <1|2|3>` arguments were specified, execute all validation tests specified in parser argument.
7. Return processed workbook(s) to users.
8. Display a terminal summary of validation results and any failed file runs.

# Database_Architecturing_Documentation_Tool

This *validator tool* supports metadata auditing efforts by automating the validation of documented metadata against the corresponding data available in Dremio. Its primary purpose is to identify inconsistencies between the documentation and the corresponding Dremio records, reducing the effort required for manual validation. These *inconsistencies/errors* are highlighted in the file's output to display exactly where each discrepancy is, and a specified color that corresponds to the type of flag being raised. The tool primarily operates on the *Full Mapping* worksheet, although its configuration-driven design allows for adaptability to evolving business requirements and alternative Excel workbook formats.

### Error Classification Matrix
| Error Code Group | Flag Type | Color |
|------------------|-----------|-------|
| `100` errors | Blank fields, exception of `N/A` values | 🔴 Red |
| `200` errors | Source-to-Target mismatch | 🟣 Purple |
| `300` errors | Target-to-Dremio mismatch | 🔵 Blue |
| -c f | Colorblind argument that mutes colors | ⚫ Grey |

<br><br>

# Parser Arguments
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
Additionally, users who prefer to specify where to receive all validated files can use the `-o`/`--output`. Files can be received in the same folder as the argument for ingestion, `-f`/`--filepath`. All output files are marked with *RESULTS_X_* prefix to distinguish from older versions. The *X* is a counter that continuously increasing every time the output already has the prefix *RESULTS_*. To perform a test, users must specify the `-t`/`--test` argument, with any combination of the three available options `1|2|3`. On the contrary, users do not have to run any test, in this case the script will only utilize the `formatter()` function which is configurable. 

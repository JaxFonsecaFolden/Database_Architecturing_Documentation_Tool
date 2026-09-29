# Validation Configuration & Customization

This script is configuration-oriented as a way to adapt to business and team specific needs. Configuration files are drafted as `*.tomle` files and converted into dictionaries or list, which is utilized throughout the entirety of the *Validator Tool*. Further explanation of each configuration and how to customize is displayed below. 

## Mapping Configuraton (`Error1`, `Error2`, `Error3`)

```toml
[Error1]
101 = "column1"
102 = "column2"
...
190 = "column regarding the source data type"
191 = "column regarding the size/length of source data type"
192 = "column regarding the regarding the time zone"
```

These mappings determine which validations are executed and which columns participate in each rule. The validator is intentionally configuration-driven, allowing validations to be enabled, disabled, or reassigned without mopdifying the underlying validation logic. __To disable a specific validation__, remove or comment out its corresponding mapping entry. Any rule not present in its associated error map will be skipped during execution. The mappings also define the relationships between source, target, and Dremio metadata. __When modifying worksheet schemas or renaming columns__, the associated error mappings and validation methods should be reviewed and updated to ensure validations continue referencing the correct columns.

*Relevant Code:*

- [Config.toml](Config.toml)
- [ErrorCodes](Tool_Lib/ErrorCodes.py)

## Worksheet Columns/Order

```toml
...
```

This script requires all ingested metadata to follow a defined worksheet schema, which is derived directly from `ERROR1_MAP`. As the schema source of truth, the insertion order of `ERROR1_MAP` determines the final column arrangement of the validated worksheet. __When adding, removing, renaming, or reordering columns__, corresponding updates to `_map_header_colors()` may be required to keep workbook formatting aligned with the column headers. The same schema is referenced by `error!()` in *ErrorCodes.py*, where several special-case validation rules depend on specific column names. __When modying existing header names__, the associated validation logic and error mappings must also be reviewed and updated to avoid inaccurate validation results. 

*Relevant Code:*
- [Config.toml](Config.toml) `ERROR1_MAP`, `EXPECTED_COLUMNS`, `_map_header_colors()`
- [ProcessMetadata.py](Tool_Lib/ProcessMetadata.py) `def check_columns()`
- [ErrorCodees.py](Tool_Lib/ErrorCodes.py) `error1()`

## Worksheet Selection (`map_sheet`)

```python
# Config.toml
Metadata_Worksheet = "Full Mapping"

# utilities.py Variable
WORKSHEET = config["Metadat_Worksheet"]

# Constructor
def __init__(
    self, 
    filepath: str,
    output_file: str,
    map_sheet: Optional[str] = WORKSHEET
):

# Created Object
run = ProcessMetadata(
    filepath=file,
    output_file=results_file,
    map_sheet=...
)
```

By default, the validtor targets the *Full Mapping* worksheet. To validate a different a different worksheet, there are 3 options dependent on TECH or Business needs. __(2)__ Update the configuration file to have the expected sheet name. __(2)__ Can update the constructor to have specified fall-back option for worksheet name, however this is not recommended as the variable explicitly points directly to the configuration file. __(3)__ When calling the `ProcessMetadata()` object in the *Main.py* file, provide a sheet name in the `map_sheet=` argument.

*Relevant Code:*
- [Config.toml](Config.toml) `Metadata_Worksheet`
- [Utilities.py](Tool_Lib/Utilities.py) `WORKSHEET`
- [ProcessMetadata.py](Tool_Lib/ProcessMetadata.py) `ProcessMetadata.__init__()`
- [Main.py](Main.py) `# ---2. Ingest Mapping Docs---`

## Source-to-Target Table Naming `TABLE_PREFIX`
```python
# Logic
TABLE_PREFIX = "eif_"
expected = (TABLE_PREFIX + self.metadata[source_col])

# Prefix Ex.
TABLE_PREFIX = ""
TABLE_PREFIX = "prefix_"
```

In the current implementation, source tables are expected to map to target views prefixed with `eif_`. If your team's naming convention differes, update `TABLE_PREFIX` accordingly. If no prefix is required, set the value to an empty string(`""`). If table naming conventions cannot be consistently enforced, validation rule `201` can be disabled, refer to [Mapping Configuration](#mapping-configuraton-error1-error2-error3).

*Relevant Code:*
- [Config.toml](Config.toml) `Table_Prefix`
- [Utilities.py](Tool_Lib/Utilities.py) `TABLE_PREFIX`
- [ErrorCodes.py](Tool_Lib/ErrorCodes.py) `def error2() -> Error 201`

## Data Tye Mapping (`type_conversion`)

```toml
```

This mapping defines which source-system data types are considered valid conversions to each Dremio data type and is used by validation rule `203`. __When onboarding new source systems__, update `type_conversion` to include any additional data type representations. Data type conversions are not always absolute, so `error 203` should be treated as a logical validation rather than a strict enforcement rule.

*Relevant Code:*
- [Config.toml](Config.toml) `type_conversion`
- [ErrorCodes.py](Tool_Lib/ErrorCodes.py) `def error2() -> Error 203`

<br><br>

# Excel Workbook Schema Configuration

## Header Colors (`HeaderColors_Hex`)

```toml
```

The workbook's header colors are defined in `HeaderColros_Hex`, which stores all color values as hexadecimal codes required by `OpenPyXL`. The `__map_header_colros()` method assigns those colors to columns using positional ranges within `expected_columns`. __When reordering, adding, or removing columns__, `__map_header_colors()` will need to be updated appropriately to ensure headers remain group under the intended color sheme. Changes made to `HeaderColros_Hex` only affect workbook presentation and do not impact validation logic. 

*Relevant Code:*
- [Config.toml](Config.toml) `HeaderColros_Hex`
- [Utilities.py](Tool_Lib/Utilities.py) `_map_header_colors()`, `HEADER_COLORS`
- [Formatter.py](Tool_Lib/Formatter.py) `def format() -> ---Header Coloring Convention---`

### Flag Fill Colors (`ErrorColors`)

```toml
```

This dictionary defines the fill color and legend description associated with each error category. The configured colrs are used when highlighting validation findings in the worksheet, while descriptions are used to generate the legend in the *Extra Errors* sheet. __Modifying color values__ only affects workbook presentation and does not impact validation logic or error detection. *Developer recommendation* is to keep colors distinguishable and separate of each error grouping.

*Relevant Code:*
- [Config.toml](Config.toml) `ErrorColors`
- [Utilities.py](Tool_Lib/Utilities.py) `ERROR_COLORS`
- [Formatter.py](Tool_Lib/Formatter.py) `def format() -> ---ErrorCodes fill colors---`

## Table styleing (`EXTRA_ERRORS_STYLE`, `METADATA_STYLE`)

```toml
```

These variables define the Excel table styles applied to th e*Validation Results* and *Extra Errors* worksheets. While customization is fully supported, the current styles are recommended because they provide clear visibility of validation highlights and workbook formatting. The style number does not represent a fixed color, as its appearance dpeends on the workbook theme and OpenPyXL's rendering. If no style is specified, the script defaults to TableStyleLight1.

*Relevant Code:*
- [Config.toml](Config.toml) `TableStyles`
- [Utilities.py](Tool_Lib/Utilities.py) `TABLE_STYLES`
- [Formatter.py](Tool_Lib/Formatter.py) `def create_table()`

## Field Styling (`EXTRA_ERRORS_STYLE`, `METADATA_STYLE`)

```toml
[ColumnStyles]
Upper = ["column_X", "column_Y", "column_Z"]
Title = ["column_X", "column_Y", "column_Z"]
```

These variables define which fields need to be uppercased and which fields are considered titles. Note *N/A* will always be uppercased and that all other fields not specified will be lowercase.

*Relevant Code:*
- [Config.toml](Config.toml) `TableStyles`
- [Utilities.py](Tool_Lib/Utilities.py) `\TABLE_STYLES`
- [Formatter.py](Tool_Lib/ProcessMetadata.py) `def replace_worksheet()`
import re
import tomllib
from pathlib import Path

def generate_output_name(
    filepath: str,
    output_dir: Path
) -> str:
    """__Utility Method__.
    Generates output filename with prefix *RESULTS_X_* based on given current filename.

    Args:
        filepath (str): 
            Path to file being ingested
        output_dir (Path): 
            Directory to file which th efile will be exported

    Returns:
        string:
            Corrected *RESULTS_X_* versioning of ingested file
    """
    stem = Path(filepath).stem
    match = re.match(r"*Results_(\d*)_(.+)$", stem)
    if match:
        current_num = int(match.group(1))
        base_name = match.group(2)
        next_num = current_num + 1
    else:
        base_name = stem
        next_num = 1
    
    return output_dir / f"RESULTS_{next_num}_{base_name}.xlsx"

# ---Open Config File---
library_dir = Path(__file__).parent
config_path = library_dir / ".." / "Config.toml"

with open(config_path, "rb") as f:
    config = tomllib.load(f)
    
def _int_keys(dictionary: dict[str, str]) -> dict:
    """__Private Method__.
    Standardized dictionary with key values as Integer types

    Args:
        dictionary (dict[str,str]) : Configuration variable containing the dictionary

    Returns:
        dict[int,str] : Standardized dictionary with key values as Integer types
    """
    return {int(k): v for k, v in dictionary.items()}

# ---Function Utilized in Formatter.py---
def _combine_error_maps() -> dict:
    """__Private Method__.
    Combines all error_maps.

    Returns:
        dict[int,list[str]] : *Key* is the error code and the corresponding
                              *Value* is a list of all columns associated with the error code
    """
    highlights = {}
    for code, col in ERROR1_MAP.items():
        # Code Automatically Updates length if necessary
        if col == "length":
            continue
        highlights[code] = [col]
    for code, col in ERROR2_MAP.items():
        highlights[code] = list(col.values())
    for code, col in ERROR3_MAP.items():
        highlights[code] = list(col.values())
    return highlights

def _map_header_colors() -> dict:
    """__Private Method__.
    Maps the headers to their preferred colors.
    *Column to Color is hardcoded.*

    Returns:
        dict[str,hex] : *Key* is th ecolumn header name and
                        *Value* is the hex color
    """
    header_colors = {}
    header = config["HeaderColors_Hex"]
    for col in EXPECTED_COLUMNS[0:4]:
        header_colors[col] = header["ORANGE"]
    for col in EXPECTED_COLUMNS[4:8]:
        header_colors[col] = header["GREY"]
    for col in EXPECTED_COLUMNS[8:13]:
        header_colors[col] = header["DARKORANGE"]
        header_colors["last updated comments"] = config["HeaderColors_Hex"]["Grey"]
    for col in EXPECTED_COLUMNS[14:22]:
        header_colors[col] = header["BLUE"]
    for col in EXPECTED_COLUMNS[22:]:
        header_colors[col] = header["ORANGE"]
    return header_colors

# ---Configuration Variables---
TABLE_PREFIX = config["Table_Prefix"]
WORKSHEET = config["Metadata_Worksheet"]
TYPE_CONVERSION = config["Type_Conversions"]
ERROR1_MAP = _int_keys(config["Error1"])
ERROR2_MAP = _int_keys(config["Error2"])
ERROR3_MAP = _int_keys(config["Error3"])
ALL_ERRORS = _combine_error_maps()
EXPECTED_COLUMNS = list(ERROR1_MAP.values())
HEADER_COLORS = _map_header_colors()
ERROR_COLORS = _int_keys(config["ErrorColors"])
TABLE_STYLES = config["TableStyles"]
COLUMN_STYLES = config["ColumnStyles"]
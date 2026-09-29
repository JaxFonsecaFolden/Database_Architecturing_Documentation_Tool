import pandas as pd 

from Tool_Lib.Utilities import (
    TABLE_PREFIX, TYPE_CONVERSION,
    ERROR1_MAP, ERROR2_MAP, ERROR3_MAP
)

class ErrorCodes:
    """
    Provides personalized flag raising utilities in the metadata ingested 
    based on business rules, and used by *ProcessMetadata*. 
    """
    def _add_error_code(self, mask: pd.Series, error_code: int | str) -> None: 
        """**Private Helper Method.** 
        Adds error codes to the specified rows, alongside existing error codes. 

        Args: 
            mask (pd.Series): A boolean mask identifying target rows. 
            error_code (int | str): Error code to append.
        """
        # Ensure ErrorCodes column exists
        if 'ErrorCodes' not in self.metadata.columns:
            self.metadata['ErrorCodes'] = ""

        # Safely append error codes without leading/trailing commas
        self.metadata.loc[mask, 'ErrorCodes'] = (
            self.metadata.loc[mask, 'ErrorCodes']
            .fillna("")
            .astype(str) 
            .apply(lambda x: f"{x}, {error_code}".strip(", ") if x and x != "nan" else str(error_code))
        )

    def error1(self) -> None:
        """**Null Value Error.** 
        Contains special rules for *Length* and *Time Zone*. 
        """
        for error_code, col in ERROR1_MAP.items():
            if error_code == 192: 
                datatype_mask = (
                    self.metadata["source data type"]
                    .str.contains(r"TIME|MI:SS", case=False, na=False)
                )
                timezone_null_mask = self.metadata["time zone"].isna() 
                mask = datatype_mask & timezone_null_mask 

            # --- Raises & Changes if a length is present in datatype --- 
            elif error_code == 191: 
                # Extract content inside parenthesis, e.g. VARCHAR(50) -> 50
                datatype_precision = (
                    self.metadata["source data type"]
                    .str.extract(r"\(([^)]+)\)")[0]
                )
                # --- .isna() checks: np.nan, None, pd.na, NaT ---
                mask = (
                    datatype_precision.notna() 
                    & (
                        self.metadata["length"].isna() 
                        | (datatype_precision != self.metadata["length"].astype(str))
                    )
                ) 
                self._add_error_code(mask, error_code)
                self.metadata.loc[mask, "length"] = datatype_precision[mask]

            # --- Data type can only be flagged if column/field name is present --- 
            elif error_code == 190: 
                source_col_present = (
                    self.metadata["column/field name"]
                    .notna()
                ) 
                mask = (
                    self.metadata["source data type"].isna() 
                    & ~source_col_present 
                )

            # --- Flags all other columns --- 
            else: 
                mask = self.metadata[col].isna() 


            self._add_error_code(mask, error_code)

    def error2(self): 
        """**Mismatch Error.**
        Contains special rules dependent on mapping rule, 
        and Source/Target column matching. 
        """
        for error_code, nested_dict in ERROR2_MAP.items():
            source_col = nested_dict["source_column"]
            target_col = nested_dict["target_column"]
            if error_code == 201: 
                # --- For teams with Table/View Naming Convention --- 
                expected = (TABLE_PREFIX + self.metadata[source_col])
                valid_rows = (
                    self.metadata[source_col].notna()
                    & self.metadata[target_col].notna() 
                    & ~self.metadata[source_col].isin(["n/a"])
                    & ~self.metadata[target_col].isin(["n/a"])
                )
                mask = (
                    (valid_rows)
                    & 
                    (self.metadata[target_col] != expected)
                )
                self._add_error_code(mask, error_code)
            # --- Field Name Checker --- 
            elif error_code == 202: 
                # Flags only if mapping rule is empty or straight move --- 
                applicable_rows = (
                    self.metadata["mapping rule"].eq("straight move")
                    | self.metadata["mapping rule"].isna() 
                )
                source_values = self.metadata[source_col]
                target_values = self.metadata[target_col]
                mask = (
                    applicable_rows
                    & (source_values != target_values)
                )
                self._add_error_code(mask, error_code)
            # --- Source/Target Datatype Conversion --- 
            elif error_code in (203, 204): 
                source_values = (
                    self.metadata[source_col]
                    .str.replace(r"\(.*?\)", "", regex=True)
                )
                target_values = self.metadata[target_col]
                # --- Nullifies when source is "N/A" --- 
                mask = (
                    (
                        source_values.isna()
                        & target_values.notna() 
                    )
                    | 
                    (
                        source_values.isna() 
                        & ~source_values.isin(["n/a"])
                        & target_values.isin(["n/a"])
                    )
                    |
                    (
                        source_values.isin(["n/a"])
                        & target_values.isna() 
                    )
                )
                # --- Special Business Rule --- 
                for schema, valid_types in TYPE_CONVERSION.items(): 
                    # --- Source containing "date" in value --- 
                    if schema == "date": 
                        date_mask = (
                            target_values.eq("date")
                            & ~source_values.str.contains(
                                "date|yy-mm-ddbhh", 
                                na=False
                            )
                        )
                        mask = mask | date_mask 
                        continue 
                    # --- Source containing "time" in value -- 
                    elif schema == "timestamp": 
                        timestamp_mask = (
                            target_values.eq("timestamp")
                            & ~source_values.str.contains(
                                "time|mi:ss", 
                                na=False
                            )
                        )
                        mask = mask | timestamp_mask
                        continue 
                    # --- All other datatypes follow dictionary conversions stated --- 
                    else: 
                        generic_mask = (
                            target_values.eq(schema)
                            & ~source_values.isin(valid_types)
                        )
                        mask = mask | generic_mask
                self._add_error_code(mask, error_code)

    def error3(self): 
        """**Dremio Mismatch Error.** 
        Contains special case handling for missing columns
        """
        for error_code, nested_dict in ERROR3_MAP.items(): 
            # --- Target vs Dremio view name --- 
            if error_code == 301: 
                for path, df in self.dremio.items(): 
                    if df is None: 
                        mask = self.metadata["sql_path"] == path
                        self._add_error_code(mask, error_code)
            # --- Target vs Dremio column name --- 
            elif error_code == 302: 
                map_header = nested_dict["target_column"]
                dremio_header = nested_dict["dremio_column"]
                for path, df in self.dremio.items(): 
                    if df is None: 
                        continue 
                    map_view = self.metadata[self.metadata["sql_path"] == path]
                    map_cols = set(map_view[map_header].dropna())
                    dremio_cols = set(df[dremio_header])
                    extra_mapping_cols = map_cols - dremio_cols 
                    # --- Raise flag for any extra/mis-named columns --- 
                    for col in extra_mapping_cols: 
                        mask = (
                            (self.metadata["sql_path"] == path)
                            & 
                            (self.metadata[map_header] == col)
                        )
                        self._add_error_code(mask, error_code)
                    # --- Setting self.data_issues, if any appliable --- 
                    missing_mapping_cols = [
                        f"Missing: {col}" for col in (dremio_cols - map_cols)
                    ]
                    if missing_mapping_cols: 
                        self.data_issues[path] = list(missing_mapping_cols)
            # Target vs Dremio Schema 
            elif error_code == 303: 
                map_header = nested_dict["target_column"]
                dremio_header = nested_dict["dremio_column"]
                for path, df in self.dremio.items():
                    if df is None: 
                        continue 
                    map_view = self.metadata[self.metadata["sql_path"] == path]
                    # --- Comparison only comprises of != pd.NA or != "N/A" values 
                    map_cols = set(map_view["target column name"].dropna())
                    dremio_cols = set(df["column_name"])
                    common_cols = map_cols & dremio_cols
                    for col_name in common_cols: 
                        map_type = map_view.loc[map_view["target column name"] == col_name, map_header].iloc[0]
                        dremio_type = df.loc[df["colum name"] == col_name, dremio_header].iloc[0]
                        if pd.isna(map_type) or pd.isna(dremio_type): 
                            continue 
                        # --- Raises flag if self.metadata does not a field present in Dremio view --- 
                        if map_type != dremio_type: 
                            mask = (
                                (self.metadata["sql_path"] == path)
                                & 
                                (self.metadata["target column name"] == col_name)
                            )
                            self._add_error_code(mask, error_code)
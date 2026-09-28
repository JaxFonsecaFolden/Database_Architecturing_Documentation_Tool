from typing import Optional
from pathlib import Path
from openpyxl.utils.dataframe import dataframe_to_rows
import openpyxl
import warnings
import pandas as pd

from Tool_Lib.Formatter import Formatter
from Tool_Lib.ErrorCodes import ErrorCodes
from Tool_Config.Utilities import EXPECTED_COLUMNS, COLUMN_STYLES, WORKSHEET

class ProcessMetadata(Formatter, ErrorCodes):
    """
    Handles validation, standardization, return of mapping documents.
    Inherits ErrorCodes and Formatter

    Attributes:
        filepath (str): 
            Path to mapping file(s)
        output_file (str): 
            File(s) output naming convention
        map_sheet (opt | str): 
            Specific sheet being ingested
        metadata (pd.DataFrame): 
            Normalized mapping document
        dremio (dict[str, pd.DataFrame]): 
            Nested dremio views and metadata
        missing (dict[str, pd.DataFrame]): 
            Dictionary of missing columns within each document ingested
        bad_path (dict[str, list[str]]): 
            Dictionary of bad SQL paths within each document ingested
        col_missing (str): 
            String containing a list of columns missing
        col_extra (str): 
            String containing a list of columns not specified in *EXPECTED_COLUMNS*
        colorblind (bool): 
            Boolean containing the trigger from parser argument flagging all errors in greyscale
    """
    
    def __init__(
        self, 
        filepath: str, 
        output_file: str, 
        map_sheet: Optional[str] = WORKSHEET
    ):
        """
        Initialize MappingDocument Object,
        handles standardization and catches rule based errors.

        Args:
            filepath (str): 
                Path to th emapping file(s)
            output_file (str): 
                Naming convention of file(s) pushed out
            map_sheet (opt | str):
                Excel sheet containing metadata
        """
        # ---Args---
        self.map_sheet = map_sheet
        self.output_file = output_file
        self.filepath = filepath
        # ---Attributes---
        self.metadata = self._create_df()
        if "ErrorCodes" in self.metadata.columns.tolist():
            self.metadata = self.metadata.drop(columns=["ErrorCodes"])
        self.new_sheet = None
        self.workbook = openpyxl.load_workbook(self.filepath)
        self.dremio = {}
        self.data_issues = {}
        self.bad_path = {}
        self.cols_missing = []
        self.cols_extra = []
        self.colorblind = []
        
    def _create_df(self) -> pd.DataFrame:
        """__Private Method__
        Initializing and normalizing all datasets.

        Returns:
            pd.DataFrame: 
                Normalized mapping document
        """
        data = pd.read_excel(
            self.filepath,
            sheet_name=self.map_sheet,
            keep_default_na=False
        )
        # ---Columns---
        data.columns = (data.columns
            .str.replace(r'[\r\n]+', ' ', regex=True)
            .str.replace(r'\s+', ' ', regex=True)
            .str.strip()
            .str.lower()
            .astype("string")
        )
        # ---Metadata---
        string_cols = data.columns
        data[string_cols] = data[string_cols].apply(lambda col: col
            .astype("string")
            .str.replace(r'\s+', ' ', regex=True)
            .str.strip()
            .str.lower()
            # ---All empty cells as pd.NA variable---
            .replace("", pd.NA)
            # ---Standardize n/a---
            .replace(["na", "n.a.", "n.a", "na."], "n/a")
        )
        data = data.drop_duplicates()
        return data
    
    def check_columns(self) -> bool:
        """
        Checks for all columns that are expected,
        creates teh official SQL path in each document,
        and organizes all headers appropriately.

        Returns:
            boolean: 
                *True* if the column structure is valid,
                otherwise *False*
        """
        current_columns = self.metadata.columns.tolist()
        self.cols_missing = [
            col for col in EXPECTED_COLUMNS
            if col not in current_columns
        ]
        self.cols_extra = [
            col for col in current_columns
            if col not in EXPECTED_COLUMNS
        ]
        # ---Organizing Headers---
        if not self.cols_missing and not self.cols_extra:
            self.metadata = self.metadata[EXPECTED_COLUMNS].copy()
            self.metadata["sql_path"] = (
                self.metadata["dremio view path"] + "." + self.metadata["target table/file name"].fillna("")
            )
            self.metadata["ErrorCodes"] = ""
            return True
        return False
        
    def set_colorblind(
        self, 
        parser_arg: list
    ):
        """
        Set's all error flags to grey scale.
        *Does not affect workbook schema.*

        Args:
            parser_arg (list):
                *-c*, *--colorblind* argument from *Main* file
        """
        if parser_arg:
            self.colorblind = parser_arg
            
    def set_dremio(
        self,
        dremio_output: dict[str, pd.DataFrame]
    ) -> bool:
        """
        Processes and normalize the pre-made dictionary of the Dremio DESCRIBE views.
        Then setting the attribute *self.dremio*

        Args:
            dremio_output (dict[str, pd.DataFrame]): 
                Dremio views output pulled from *Main* script

        Returns:
            boolean:
                *True* if the SQL path given was successful and view was ingested,
                otherwise *False*
        """
        self.dremio = dremio_output
        # ---Match Dremio SQL Paths---
        for path, df in dremio_output.items():
            if df is None:
                continue
            # ---Standardize Views---
            df = df[["COLUMN_NAME", "DATA_TYPE", "IS_NULLABLE"]]
            df.columns = df.columns.str.lower()
            for col in df.columns:
                if pd.api.types.is_string_dtype(df[col]):
                    df[col] = df[col].str.lower()
            self.dremio[path] = df
        return bool(self.dremio)
    
    def set_issues(
        self,
        issues: dict[str, list[str]]
    ) -> bool:
        """
        Processess and normalize the pre-made dictionary of failed Dremio paths.
        Then setting the attribute *self.dremio*.

        Args:
            issues (dict[str, list[str]]): 
                *Key*, *Value* pair of failures, or problematic data

        Returns:
            boolean: 
                *True* if the SQL path given was successful and 
                paths were ingested, otherwise *False*
        """
        for path, error_message in issues.items():
            self.data_issues[path] = error_message
        return bool(self.data_issues)
        
    def replace_worksheet(self):
        """
        Copy and replace processed worksheet with newly prepared one.
        """
        # ---Prepare Output Info---
        self.metadata = self.metadata.drop("sql_path", axis=1)
        self.metadata = self.metadata.fillna("")
        title_cols = COLUMN_STYLES["Title"]
        upper_cols = COLUMN_STYLES["Upper"]
        self.metadata[title_cols] = self.metadata[title_cols].apply(lambda col: col.str.title())
        self.metadata[upper_cols] = self.metadata[upper_cols].apply(lambda col: col.str.title())
        self.metadata = self.metadata.replace("n/a", "N/A")
        # ---Remove & Create New Worksheet---
        ws = self.workbook[self.map_sheet]
        self.workbook.remove(ws)
        ws = self.workbook.create_sheet(self.map_sheet)
        # ---Copy Dataframe Info---
        for row in dataframe_to_rows(
            self.metadata,
            index=False,
            header=True
        ):
            ws.append(row)
            
    def export_workbook(
        self, 
        output_name: str
    ) -> bool:
        """
        Export current Excel Workbook.

        Args:
            output_name (str): 
                Generated output file naming convention
                
        Returns:
            boolean: *True* if executed, otherwise *False*
        """
        try:
            with warnings.catch_warnings():
                warnings.filterwarnings(
                    "ignore",
                    message="File may not be readable"
                )
                self.workbook.save(output_name)
            self.workbook.close()
            return True
        except Exception as e:
            print(e)
            return False
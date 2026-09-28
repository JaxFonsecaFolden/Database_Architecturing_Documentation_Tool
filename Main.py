import argparse
import sys
import getpass
import pandas as pd
from pathlib import Path
from tqdm import tqdm

# Personal Libraries
from Tool_lib.DremioConnection import DremioClient
from Tool_lib.ProcessMetadata import ProcessMetadata
from Tool_Config.ColorMaps import (
    ERROR_COLORS, generate_output_name
)

def main():
    
    # 0. ---Parser Arguments---
    parser = argparse.ArgumentParser(description="Compare and Contrast Dremio Column Details")
    # files
    parser.addargument("-f", "--filepath", required=True, help="Path to maping document")
    parser.add_argument("-o", "--output", default=None, help="Output folder for the result")
    # dremio
    parser.add_argument("-u", "--user", default=None, help="Dremio Username")
    parser.add_argument("-e", "--env", default="uat", help="Dremio environment (uat | prod)")
    # error codes
    parser.add_argument("-t", "--test", type=str, deafult=[], help="Testing different error runs (1,2,3)")
    # other args
    parser.add_argument("-c", "--colorblind", default=[], help="Enable colorblind-freindly output (t=terminal, f=file, tf=both)")
    parser.add_argument("-d", "--details", defualt=["c", "d", "f"], help="Terminal details for file validation notifications (c=column chedcker, d=dremio views, f=failed files)")
    args = parser.parse_args()
    
    # Terminal Formatting
    BOLD = "\033[1m"
    RESET = "\033[0m"
    
    # Color Arg
    BLUE1 = "\033[34m"
    BLUE2 = "\033[94m"
    RED = "\033[91m"
    PURPLE = "\033[95m"
    
    if args.colorblind:
        if isinstance(args.colorblind, str):
            args.colorblind = list(
                args.colorblind.lower().replace(",", "").strip()
            )
        valid_color_args = {"t", "f"}
        if not set(args.colorblind).issubset(valid_color_args):
            raise ValueError(
                "Invalid colorblind option(s): t=terminal, f=file, tf= both"
            )
    if "t" in args.colorblind:
        BLUE1 = BLUE2 = RED = PURPLE = ""
        
    # Details Error
    if args.details:
        if isinstance(args.details, str):
            args.details = list(
                args.details.lower().replace(",", "").strip()
            )
        valid_details_args = {"c", "d", "f"}
        if not set(args.details).issubset(valid_details_args):
            raise ValueError(
                "Invalid details option(s): c=column checker, d=dremio views, f=failed files"
            )
            
    # Error Arg
    error_options = {
        str(code): details["description"]
        for code, details in ERROR_COLORS.items()
    }
    if args.test:
        if isinstance(args.test, str):
            args.test = list(
                args.test.lower().replace(",", "").strip()
            )
        valid_test_args = set(error_options.keys())
        if not set(args.test).issubset(valid_test_args):
            raise ValueError(
                "Invalid test option(s): 1, 2, 3"
            )
            
    # Credentials
    if "3" in args.test:
        username = args.user or input(f"{bool}Username:{RESET}  ")
        password = getpass.getpass(f"{BOLD}Password:{RESET}  ")
        
        # Env Arg
        host = {
            "uat": "...Connection Str...",
            "prod": "...Connection Str..."
        }.get(args.env.lower())
        if not host:
            raise ValueError('Must specify "uat" or "prod"')
        
    # File Args
    files_failed = {}
    files_to_process = []
    input_path = Path(args.filepath)
    if input_path.is_file():
        if input_path.suffix.lower() not in (".xlsx", ".xls"):
            files_failed[input_path.name] = "Wrong file type"
        else:
            files_to_process.append(input_path)
    elif input_path.is_dir():
        for file in input_path.iterdir():
            if not file.is_file():
                continue
            if file.suffix.lower() not in (".xlsx", ".xls"):
                files_failed[file.name] = "Wrong file type"
                continue
            files_to_process.append(file)
    else:
        sys.exit(f"\nInvalid path:  {args.filepath}\n")
    if not files_to_process:
        sys.exit("\nNot valid Excel file found. Exiting\n")
    
    # Output Arg
    output_dir = Path(args.output) if args.output else Path.cwd()
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # For terminal spacing
    print()
    
    # 1. ---Test Dremio Connection---
    if "3" in args.test:
        try:
            dremio_credentials = DremioClient.connect(
                host=host,
                port=...,
                username=username,
                password=password
            )
        except Exception as e:
            print(f"{BOLD}{BLUE1}-----Failed to connect to DREMIO-----{RESET}")
            print(f"{e}")
            sys.exit("\nTest connection failed. Exiting\n")
            
    # 2. ---Ingesting Mapping Doc---
    files_num = len(files_to_process)
    files_ingested = 0
    print(f"{BOLD}{BLUE1}-----Checking Mapping Documents-----{RESET}")
    # Progress Bar for Files
    pbar = tqdm(files_to_process, desc="Ingest File(s)", unit="File")
    for file in pbar:
        pbar.set_description(f"{BOLD}Running{RESET} {file.name}")
        results_file = generate_output_name(file, output_dir)
        try:
            run = ProcessMetadata(file, results_file)
            # Colorblind Arg
            if args.colorblind:
                run.set_colorblind(args.colorblind)
            # Column Checker
            if not run.check_columns():
                files_failed[file.name] = "Failed column validation"
                if "c" in args.details:
                    tqdm.write(f"\n{BOLD}> {file.name}{RESET}")
                    if run.cols_missing:
                        tqdm.write(f"   Missing Columns: {run.cols_missing}")
                    if run.cols_extra:
                        tqdm.write(f"   Extra Columns: {run.cols_extra}")
                    files_failed[file.name] = "Failed column validation"
                continue
            files_ingested += 1
        except Exception as e:
            files_failed[file.name] = str(e)
            continue
        
    # 3. ---Ingest Dremio Views---
        missing_sql_path = 0
        missing_view_name = 0
        bad_path = []
        if "3" in args.test:
            doc_sql_paths = run.metadata["sql_path"].unique().tolist()
            dremio_views = {}
            dremio_bad_paths = {}
            # Progress Bar for Dremio Views
            for dremio_path in tqdm(
                doc_sql_paths,
                desc=f"{BOLD}Dremio Views{RESET}",
                unit="View",
                leave=False
            ):
                try:
                    dremio_views[dremio_path] = dremio_credentials.describe(dremio_path)
                except Exception as e:
                    # Failed Path/View Error Catch
                    if pd.isna(dremio_path):
                        error_msg = f"Check target view name and dremio path"
                        key = "SQL path is completely missing"
                        missing_sql_path += 1
                    else:
                        view_name = dremio_path.split(".")[-1]
                        if not view_name:
                            error_msg = "Missing view name"
                            missing_view_name += 1
                        else:
                            error_msg= f"Bad Path:{view_name}"
                            bad_path.append(view_name)
                        key = dremio_path
                    # Apply to dictionaries
                    dremio_bad_paths[key] = [error_msg]
                    dremio_views[dremio_path] = None
                    
            if bad_path and ("d" in args.details):
                tqdm.write(f"\n{BOLD}> {file.name}{RESET}")
                tqdm.write(f"   Bad Paths: {bad_path}")
                
            # Set MappingDoc Attribute
            if dremio_bad_paths:
                run.set_dremio(dremio_bad_paths)
            if dremio_views:
                run.set_dremio(dremio_views)
                
        # 4. ---Running ErrorCodes---
        if args.test:
            try:
                for error in args.test:
                    getattr(run, f"error{error}")()
            except Exception as e:
                files_failed[file.name] = (
                    f"Failed to run error {error} > {e}"
                )
                continue
            
        # 5. ---Results---
        run.replace_worksheet()
        run.format()
        if not run.export_workbook(results_file):
            files_ingested -= 1
            files_failed[file.name] = "Failed to save/export to excel workbook"
            continue
        
    # 6. ---Terminal Summary---
    env_arg = f" ({args.env.upper()})" if "3" in args.test else ""
    print(f"\n{BOLD}{BLUE1}-----Running Test{env_arg}-----{RESET}")
    if args.test:
        for error in args.test:
            if error == "1":
                print(f"{RED}[{error}]{RESET} {error_options[error]}")
            elif error == "2":
                print(f"{PURPLE}[{error}]{RESET} {error_options[error]}")
            elif error == "3":
                print(f"{BLUE2}[{error}]{RESET} {error_options[error]}")
        else:
            print("[0] No test has been specified, formatting documents only")
    
    if files_ingested != files_num:
        print(f"\n{BOLD}{BLUE1}-----Documents Ingested-----{RESET}")
        print(f"{BOLD}{files_ingested}/{files_num}{RESET}  :  File Run Rate")
    
    if files_failed and ("f" in args.details):
        print(f"\n{BOLD}{BLUE1}-----Failed Files/Summary-----{RESET}")
        for file, reason in files_failed.items():
            print(f"\n{BOLD}> {file}{RESET}")
            print(f"    {reason}")
    
    if __name__ == "__main__":
        main()
#!/usr/bin/env python3
"""Create empty portable data files for packaged desktop releases."""

import os
import sys

import core
import database


def create_starter_files(output_dir):
    output_dir = os.path.abspath(os.fspath(output_dir))
    os.makedirs(output_dir, exist_ok=True)
    database_path = os.path.join(output_dir, "BKTC_Ledger.db")
    workbook_path = os.path.join(output_dir, "BKTC_Ledger.xlsx")
    if not os.path.exists(database_path):
        database.save_database(
            database_path, core.empty_data(), reason="portable_starter", backup=False
        )
    if not os.path.exists(workbook_path):
        core.generate_xlsx(core.empty_data(), workbook_path)
    return database_path, workbook_path


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("usage: create_portable_starter.py OUTPUT_DIR")
    create_starter_files(sys.argv[1])

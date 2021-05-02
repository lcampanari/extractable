import os
import pandas as pd
import logging

logging.basicConfig(
    filename="logs.log",
    format="%(asctime)s - %(levelname)s: %(message)s",
    level=logging.DEBUG,
    datefmt="%m/%d/%Y %I:%M:%S %p",
)
logging.propagate = False


class Spreadfy:
    def __init__(
        self,
        data_directory="data-test/",
        data_extensions=(".xlsx", ".xls"),
        result_filename="result.xlsx",
        subtable_header_id=None,
        subtable_callback=None,
    ):
        self.data_directory = data_directory
        self.data_extensions = data_extensions
        self.result_filename = result_filename
        self.subtable_header_id = subtable_header_id
        self.subtable_callback = subtable_callback
        self.total_tables_extracted = 0

    def process_directory(self):
        dfs = []

        for subdir, dirs, files in os.walk(self.data_directory):
            for file in files:
                filename = os.path.join(subdir, file)

                if not self.is_file_allowed(file):
                    logging.warning(
                        f"Skipping {filename} because the file type is not one of the following: {self.data_extensions}"
                    )
                    continue

                result = self.process_subtables(filename)
                dfs.append(result)

        return self.flatten(dfs)

    def process_subtables(self, filename):
        df = self.read_file(filename)
        file_rows = df.iterrows()
        file_rows_length = len(list(df.iterrows()))
        rows = []
        tables = []

        for index, row in file_rows:
            if self.is_empty(row):
                continue

            row = row.values

            #  found subtable: save previous rows if any and start collecting rows again
            if self.is_subtable_header(row):
                if len(rows) > 0:
                    tables.append(rows)
                    rows = []
                rows.append(row)
            # only save the row if already started collecting
            elif len(rows) > 0:
                rows.append(row)
                # save last table found on file
                if index == file_rows_length - 1:
                    tables.append(rows)

        dfs = self.subtables_to_dfs(tables)

        tables_length = len(tables)
        self.total_tables_extracted += tables_length

        logging.info(f"Processed {filename}:")
        logging.info(f"  - {tables_length} table(s) extracted")

        return dfs

    def read_file(self, filename):
        df = pd.read_excel(filename, header=None)
        return df

    def subtables_to_dfs(self, tables):
        dfs = []
        for rows in tables:
            columns = rows.pop(0)
            df = pd.DataFrame(rows, columns=columns)

            if self.subtable_callback:
                df = self.subtable_callback(df)

            dfs.append(df)

        return dfs

    def is_subtable_header(self, row):
        return self.subtable_header_id in set(row)

    def is_file_allowed(self, file):
        return file.endswith(self.data_extensions)

    def is_empty(self, series):
        return series.dropna().empty

    def merge_dataframes(self, dfs, keys=None):
        return pd.concat(dfs, ignore_index=True, sort=False, keys=keys)

    def flatten(self, some_list):
        return [item for sublist in some_list for item in sublist]

    def call(self):
        dfs = self.process_directory()
        df = self.merge_dataframes(dfs)
        logging.info(
            f"Extraction completed! Total of {self.total_tables_extracted} tables extracted."
        )
        logging.info(f"Writing to {self.result_filename}")
        df.to_excel(self.result_filename)
        logging.info(f"Done!\n")

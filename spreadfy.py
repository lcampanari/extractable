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


DATA_DIR = "data/"
DATA_EXTENSIONS = (".xlsx", ".xls")

DEFAULT_FILE_NAME_RESULT = "data_result.xlsx"

FOOD_ID_COL_NAME = "FoodCode"

COL_RENAME_MAPPING = {
    "ModCode": ["ModCOode", "ModCOde", "Modcode"],
    "RawWeight": ["Raw Weight"],
    "FinalWeight": ["Final Weight"],
}


class Spreadfy:
    def __init__(
        self,
        remove_empty_rows=False,
        result_filename=DEFAULT_FILE_NAME_RESULT,
        remove_header=False,
        sub_table_header_id=None,
    ):
        self.remove_empty_rows = remove_empty_rows
        self.result_filename = result_filename
        self.remove_header = remove_header
        self.sub_table_header_id = sub_table_header_id
        self.total_tables_extracted = 0

    def process_directory(self, directory=DATA_DIR):
        dfs = []

        for subdir, dirs, files in os.walk(directory):
            for file in files:
                filename = os.path.join(subdir, file)

                if not self.is_file_allowed(file):
                    logging.warning(
                        f"Skipping {filename} because the file type is not one of the following: {DATA_EXTENSIONS}"
                    )
                    continue

                result = self.process_subtables(filename)
                # result = self.process_file(filename)

                dfs.append(result)

        return self.flatten(dfs)

    def process_file(self, filename):
        df = pd.read_excel(filename, header=None)
        df = self.process_dataframe(df)
        return df

    def process_dataframe(self, df):
        header_index = 0

        for index, row in df.iterrows():
            if self.remove_empty_rows and self.is_empty(row):
                df = df.drop(index)

        if self.remove_header:
            df = df.drop(header_index)

        return df

    def process_subtables(self, filename):
        df = pd.read_excel(filename, header=None)
        file_rows = df.iterrows()
        file_rows_length = len(list(df.iterrows()))
        rows = []
        tables = []
        dfs = []

        for index, row in file_rows:
            if self.is_empty(row):
                continue

            row = row.values

            #  found subtable: save previous rows if any and start saving again
            if self.is_subtable_header(row):
                if len(rows) > 0:
                    tables.append(rows)
                    rows = []
                rows.append(row)
            # only save the row if already started saving
            elif len(rows) > 0:
                rows.append(row)

                # in case there's only one table in the file
                if index == file_rows_length - 1:
                    tables.append(rows)

        for rows in tables:
            columns = rows.pop(0)
            df = pd.DataFrame(rows, columns=columns)

            # Fix typos on column name
            # for correct_name in COL_RENAME_MAPPING:
            #     for wrong_name in COL_RENAME_MAPPING[correct_name]:
            #         if wrong_name in df:
            #             df = df.rename(columns=dict([(wrong_name, correct_name)]))

            # copy 'id' to cell (0,0). H2O columns contain an extra row with the unit "(g)" below it
            # remove rows that contain "(g)" in the H2O column

            if "H2O" in df:
                extra_row_index = df[df["H2O"] == "(g)"].index
                if not extra_row_index.empty:
                    df.drop(extra_row_index, inplace=True)
                    index_id_missing, index_id_value = 1, 2
                else:
                    index_id_missing, index_id_value = 0, 1
            else:
                index_id_missing, index_id_value = 0, 1

            df.loc[index_id_missing, FOOD_ID_COL_NAME] = df[FOOD_ID_COL_NAME].values[
                index_id_value
            ]
            # remove empty rows
            df.dropna(how="all", axis=1, inplace=True)

            dfs.append(df)

        tables_length = len(tables)
        self.total_tables_extracted += tables_length

        logging.info(f"Processed {filename}:")
        logging.info(f"  - {tables_length} table(s) extracted")

        return dfs

    def is_subtable_header(self, row):
        return self.sub_table_header_id in set(row)

    def is_file_allowed(self, file):
        return file.endswith(DATA_EXTENSIONS)

    def is_empty(self, series):
        return series.dropna().empty

    def merge_dataframes(self, dfs, keys=None):
        return pd.concat(dfs, ignore_index=True, sort=False, keys=keys)

    def flatten(self, some_list):
        return [item for sublist in some_list for item in sublist]

    def concatenate_raw_files(self, result_filename=None):
        dfs = self.process_directory()
        df = self.merge_dataframes(dfs)
        df.to_excel(result_filename or self.result_filename)

    def call(self):
        dfs = self.process_directory()
        df = self.merge_dataframes(dfs)
        logging.info(
            f"Extraction completed! Total of {self.total_tables_extracted} tables extracted."
        )
        logging.info(f"Writing to {self.result_filename}")
        df.to_excel(self.result_filename)
        logging.info(f"Done!\n")


class SpreadfyModeRaw(Spreadfy):
    def process(self, df):
        header_index = 0

        for index, row in df.iterrows():
            if self.remove_empty_rows and self.is_empty(row):
                df = df.drop(index)

        if self.remove_header:
            df = df.drop(header_index)

        return [df]


class SpreadfyModeSubtables(Spreadfy):
    def process(self, df):
        rows = []
        tables = []
        dfs = []

        for index, row in df.iterrows():
            #  found subtable: save previous rows if any and start saving again
            if self.is_subtable_header(row):
                if len(rows) > 0:
                    tables.append(rows)
                    rows = []
                rows.append(row)
            # only save the row if already started saving
            elif len(rows) > 0:
                rows.append(row)

        for table in tables:
            columns = rows.pop(0)
            df = pd.DataFrame(rows, columns=columns)
            dfs.append(df)

        return dfs


def main():
    # spreadfy = Spreadfy()
    # spreadfy.concatenate_raw_files(result_filename="data_concatenated.xlsx")
    spreadfy = Spreadfy(
        sub_table_header_id="FoodCode", result_filename="data_final.xlsx"
    )
    spreadfy.call()


main()

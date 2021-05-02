from extractable import Extractable


def subtable_callback(df):
    # remove empty columns
    return df.dropna(how="all", axis=1)


def main():
    extractable = Extractable(
        data_directory="example/data",
        result_filename="example-result.xlsx",
        subtable_header_id="Col1",
        subtable_callback=subtable_callback,
    )
    extractable.call()


main()

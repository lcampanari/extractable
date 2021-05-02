from extractable import Extractable


def main():
    extractable = Extractable(
        data_directory="data-example/",
        result_filename="example-result.xlsx",
        subtable_header_id="Col1",
    )
    extractable.call()


main()

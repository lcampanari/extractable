# Extractable
Extract tables from multiple excel files and concatenate them into a single table.

## Usage

```python
extractable = Extractable(
    data_directory="data/",
    result_filename="result.xlsx",
    subtable_header_id="SomeColumnName",
)
extractable.call()
```
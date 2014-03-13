Parse Complex's schema from R_export.txt.xls

Usage:

    ```python
    import complex_schema
    tables = read_tables('R_export.txt.xls')
    ```

The returned data is encoded in the following structures:

Table:
- `name` - table name
- `description` - for humans
- `fields` - list of `Field`s

Field:
- `name` - field name, csv header
- `description` - for humans
- `type` - one of `char`, `num`, `date`, `prop`
- `length` - suggested length in characters

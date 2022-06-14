# odoo-JSON-RPC-VBA
The odoo's models API is easily available over JSON-RPC and accessible from the VBA language such as Excel application.

## git clone

Run the `git clone` command with the submodules.

```
git clone --recursive https://github.com/jp-rad/odoo-json-rpc-vba
```

## Tutorial workbook

Run `./tools/create_tutorial_workbook.vbs`, then open `JSON-RPC Tutorial.xlsm` and call `doAll` in the Excel VBA imidiate window.

Press `F5` key to step next.

Refer to the following document for the contents of each step.

- [odoo docs - External API](https://www.odoo.com/documentation/15.0/developer/misc/api/external_api.html)

## Blank workbook

Run `./tools/create_blank_workbook.vbs`, the `JSON-RPC Blank.xlsm` file will be created.

# Note:

## VBS Tools Runtime Error

Programmatic access to Office VBA project may be  denied.  In that case, please refer to the following page.

- [Programmatic access to Office VBA project is denied](https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)

## Date(time) Fields

When assigning a value to a Date/Datetime field, the following options are valid:

- A `date` or `datetime` object.
- A string in the proper server format:  
`YYYY-MM-DD` for Date fields,  
`YYYY-MM-DD HH:MM:SS` for Datetime fields.
- `False` or `None`.

see also [odoo docs - Date(time) Fields](https://www.odoo.com/documentation/15.0/developer/reference/backend/orm.html#date-time-fields).


The problem here is that when using `JsonConverter`, a value of type `Date` is converted to ISO format string, which the odoo server will not accept as an invalid format.
Also, VBA does not distinguish between `date` and `datetime` types, so not only `datetime` types, but even `date` types are converted to utc datetime.

To avoid these, instead of assigning `Date` type values to json, use the following conversion functions to assign `converted string`.

- `OdooJsonRpc.FormatToServerDate` for Date fields
- `OdooJsonRpc.FormatToServerUtc` for Datetime fields

Conversely, when reading from the server to the client, use the `CDate` function for Date fields and the `JsonConverter.ParseUtc` function for Datetime fields to convert them to `Date` type.

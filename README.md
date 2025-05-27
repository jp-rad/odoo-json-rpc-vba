# odoo-JSON-RPC-VBA
The odoo's models API is easily available over JSON-RPC and accessible from the VBA language such as Excel application.

## git clone

Run the `git clone --recursive` command with the submodules.

```
git clone --recursive https://github.com/jp-rad/odoo-json-rpc-vba
```

## Tutorial workbook

Run `./tools/create_tutorial_workbook.vbs`.

```
cd odoo-json-rpc-vba/tools
./create_tutorial_workbook.vbs
```

Then open `JSON-RPC Tutorial.xlsm` and call `doAll` in the Excel VBA imidiate window.

Press `F5` key to step next.

Refer to the following document for the contents of each step.

- [odoo docs - External API](https://www.odoo.com/documentation/15.0/developer/misc/api/external_api.html)

## Blank workbook

Run `./tools/create_blank_workbook.vbs`, the `JSON-RPC Blank.xlsm` file will be created.

```
cd odoo-json-rpc-vba/tools
./create_blank_workbook.vbs
```

# Note:

## VBS Tools Runtime Error

Programmatic access to Office VBA project may be  denied.  In that case, please refer to the following page.

- [Programmatic access to Office VBA project is denied](https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)

## Date(time) Fields

When working with Date or Datetime fields in Odoo from VBA, note the following:

- Odoo expects dates as strings in specific formats:
  - Date: `YYYY-MM-DD`
  - Datetime: `YYYY-MM-DDTHH:MM:SS`
- Assigning a VBA `Date` value directly to JSON will convert it to an ISO format string, which Odoo may not accept.
- VBA does not distinguish between date and datetime types, so both are converted the same way.

**To ensure compatibility:**

- Use `OdooRpc.FormatDate` to convert VBA `Date` values for Date fields.
- Use `OdooRpc.ConvertToIsoDatetime` to convert VBA `Date` values for Datetime fields.

**When reading values from Odoo:**

- Use `OdooRpc.ParseDate` to convert a string from a Date field to a VBA `Date`.
- Use `OdooRpc.ParseIsoDatetime` to convert a string from a Datetime field to a VBA `Date`.

For more details, see the [Odoo documentation on Date(time) Fields](https://www.odoo.com/documentation/15.0/developer/reference/backend/orm.html#date-time-fields).

## odoo-JSON-RPC-VBA
The odoo's models API is easily available over JSON-RPC and accessible from the VBA language such as Excel application.

- odoo-JSON-RPC-VBA (GitHub repository: [https://github.com/jp-rad/odoo-json-rpc-vba](https://github.com/jp-rad/odoo-json-rpc-vba))
- [odoo docs - External API](https://www.odoo.com/documentation/master/developer/reference/external_api.html)

### git clone

Run the `git clone --recursive` command with the submodules.

```
git clone --recursive https://github.com/jp-rad/odoo-json-rpc-vba
```

### Example workbook

Run `./create_workbook.bat`.

```
cd odoo-json-rpc-vba
./create_workbook.bat
```

**Open the Excel Files**  
- Open both `odoo-json-rpc-vba example.xlsm` and `odoo-json-rpc-vba.xlam`.

**Configure References in Visual Basic Editor (VBE)**  
- Open the Visual Basic Editor (VBE).  
- In VBE, select the project of `odoo-json-rpc-vba example.xlsm`.  
- Go to **Tools** > **References**.  
- In the References dialog, select `OdooJsonRpcVBA`.

**Execute the Tutorial Method**  
- Open the **Immediate Window** in VBE.  
- Run the following command:  

   ```vba
   DoTutorialExternalApi
   ```

Refer to the following document for the contents of each step.

- [odoo docs - External API `Calling methods`](https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods)

## Note:

### VBS Tools Runtime Error

Programmatic access to Office VBA project may be  denied.  In that case, please refer to the following page.

- [Programmatic access to Office VBA project is denied](https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)

### Date(time) Fields

When working with Date or Datetime fields in Odoo from VBA, note the following:

- Odoo expects dates as strings in specific formats:
  - Date: `YYYY-MM-DD`
  - Datetime: `YYYY-MM-DDTHH:MM:SS`
- Assigning a VBA `Date` value directly to JSON will convert it to an ISO format string, which Odoo may not accept.
- VBA does not distinguish between date and datetime types, so both are converted the same way.

**To ensure compatibility:**

- Use `OdRpc.FormatDate` to convert VBA `Date` values for Date fields.
- Use `OdRpc.ConvertToIsoDatetime` to convert VBA `Date` values for Datetime fields.

**When reading values from Odoo:**

- Use `OdRpc.ParseDate` to convert a string from a Date field to a VBA `Date`.
- Use `OdRpc.ParseIsoDatetime` to convert a string from a Datetime field to a VBA `Date`.

For more details, see the [Odoo documentation on Date(time) Fields](https://www.odoo.com/documentation/15.0/developer/reference/backend/orm.html#date-time-fields).

## CREDITS

This project utilizes the following open-source tool, which helps enhance functionality and efficiency:

- VBA-tools (GitHub repository: [https://github.com/VBA-tools](https://github.com/VBA-tools))

License: MIT License (MIT)  
Copyright (c) Tim Hall  
See the full license details in the LICENSE files or at the [VBA-tools repository](https://github.com/VBA-tools).

Support the developer:  
Tim Hall accepts donations to support his work via Patreon: [Patreon link](https://www.patreon.com/timhall)

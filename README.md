## odoo-JSON-RPC-VBA

Odoo's models API is easily accessible via JSON-RPC and can be used from VBA, such as in Excel applications.

- [Odoo Docs - External API](https://www.odoo.com/documentation/master/developer/reference/external_api.html)
- odoo-JSON-RPC-VBA (GitHub repository: [https://github.com/jp-rad/odoo-json-rpc-vba](https://github.com/jp-rad/odoo-json-rpc-vba))

### Excel VBA Example

```vb
Sub DoSearchRead()
    Dim oc As OdClient
    Dim rs As OdResult
    Dim sJson As String
    Dim n As Integer
    Dim i As Integer
    Dim sPartner As String
    
    ' OdClient
    Set oc = OdRpc.NewOdClient("https://localhost")
    oc.RefWebClient.Insecure = True
    
    ' Login
    oc.Common.Authenticate "dev_odoo", "admin", "admin"
    
    ' Search and read
    Set rs = oc.Model("res.partner").Method("search_read").ExecuteKw( _
        "[[['is_company', '=', true]]]", _
        "{'fields': ['name', 'country_id'], 'limit': 3}" _
    )
    
    ' (JSON)
    sJson = JsonConverter.ConvertToJson(rs.Result, 2)
    Debug.Print
    Debug.Print "JSON: >>>>>"
    Debug.Print sJson
    Debug.Print "<<<<<"
    Debug.Print
    
    ' (JSONLOOKUP)
    n = JSONCOUNT(sJson)
    Debug.Print "id", "country name", "(id)", "name"
    Debug.Print "--", "------------", "----", "----"
    For i = 0 To n - 1
        sPartner = JSONLOOKUP(sJson, "[" & i & "]")
        Debug.Print JSONLOOKUP(sPartner, "id"), _
                    JSONLOOKUP(sPartner, "country_id[1]"), _
                    "(" & JSONLOOKUP(sPartner, "country_id[0]") & ")", _
                    JSONLOOKUP(sPartner, "name")
    Next i
    Debug.Print
    Debug.Print "record count: " & n
    
End Sub
```

**Output**

```
JSON: >>>>>
[
  {
    "id": 15,
    "name": "Azure Interior",
    "country_id": [
      233,
      "United States"
    ]
  },
  {
    "id": 10,
    "name": "Deco Addict",
    "country_id": [
      233,
      "United States"
    ]
  },
  {
    "id": 11,
    "name": "Gemini Furniture",
    "country_id": [
      233,
      "United States"
    ]
  }
]
<<<<<

id            country name  (id)          name
--            ------------  ----          ----
 15           United States (233)         Azure Interior
 10           United States (233)         Deco Addict
 11           United States (233)         Gemini Furniture

record count: 3
```


---

### Clone the Repository

Clone the repository with submodules:

```
git clone --recursive https://github.com/jp-rad/odoo-json-rpc-vba
```

### Example Workbook

Run the batch file to create the example workbook:

```
cd odoo-json-rpc-vba
./create_workbook.bat
```

**Open the Excel Files:**  
- Open both `odoo-json-rpc-vba example.xlsm` and `odoo-json-rpc-vba.xlam`.

**Configure References in Visual Basic Editor (VBE):**  
- Open the Visual Basic Editor (VBE).  
- In VBE, select the project `odoo-json-rpc-vba example.xlsm`.  
- Go to **Tools** > **References**.  
- In the References dialog, select `OdooJsonRpcVBA`.

**Run the Tutorial Method:**  
- Open the **Immediate Window** in VBE.  
- Run the following command:  

   ```vba
   DoTutorialExternalApi
   ```

Refer to the following document for step-by-step details:

- [Odoo Docs - External API: Calling methods](https://www.odoo.com/documentation/master/developer/reference/external_api.html#calling-methods)

---

## Notes

### VBS Tools Runtime Error

Programmatic access to the Office VBA project may be denied. If this occurs, refer to:

- [Programmatic access to Office VBA project is denied](https://support.microsoft.com/en-us/topic/programmatic-access-to-office-vba-project-is-denied-960d5265-6592-9400-31bc-b2ddfb94b445)

### Date and Datetime Fields

When working with Date or Datetime fields in Odoo from VBA, note:

- Odoo expects dates as strings in specific formats:
  - Date: `YYYY-MM-DD`
  - Datetime: `YYYY-MM-DDTHH:MM:SS`
- Assigning a VBA `Date` value directly to JSON will convert it to an ISO format string, which Odoo may not accept.
- VBA does not distinguish between date and datetime types; both are converted the same way.

**To ensure compatibility:**

- Use `OdRpc.FormatDate` to convert VBA `Date` values for Date fields.
- Use `OdRpc.ConvertToIsoDatetime` to convert VBA `Date` values for Datetime fields.

**When reading values from Odoo:**

- Use `OdRpc.ParseDate` to convert a string from a Date field to a VBA `Date`.
- Use `OdRpc.ParseIsoDatetime` to convert a string from a Datetime field to a VBA `Date`.

For more details, see the [Odoo documentation on Date(time) Fields](https://www.odoo.com/documentation/15.0/developer/reference/backend/orm.html#date-time-fields).

---

## Credits

This project utilizes the following open-source tool to enhance functionality and efficiency:

- VBA-tools (GitHub repository: [https://github.com/VBA-tools](https://github.com/VBA-tools))

License: MIT License (MIT)  
Copyright (c) Tim Hall  
See the full license details in the LICENSE files or at the [VBA-tools repository](https://github.com/VBA-tools).

Support the developer:  
Tim Hall accepts donations to support his work via Patreon: [Patreon link](https://www.patreon.com/timhall)

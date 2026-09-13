# Tables and Relationships Reference

## ODBC Connection String Consistency

**All linked tables and pass-through queries in a project must use exactly the same ODBC connection string.** Mixed connection strings cause Access to open multiple ODBC connections to the same server, wasting resources and causing unpredictable authentication behaviour.

Rules:
1. The canonical connection string is defined in `db-connection.json`. All linked tables and pass-through queries must use this exact string.
2. When creating or editing a linked table definition, set `"Connect"` to the canonical string.
3. When creating or editing a pass-through query, set `dbMemo "Connect"` to the canonical string.
4. Never embed user-specific credentials (`UID=`, `PWD=`), per-user DSN variations, or protocol overrides (`Network=`) in connection strings.
5. `db-connection.json` must contain only one entry matching the canonical connection string.

## Linked Table Definitions (tbldefs/*.json)

```json
{
  "Info": {"Class": "clsDbTableDef", "Description": "Linked Table"},
  "Items": {
    "Name": "tTableName",
    "Connect": "ODBC;DSN=...;",
    "SourceTableName": "dbo.tTableName",
    "Attributes": 536870912,
    "PrimaryKey": "[ID]"
  }
}
```

- `Attributes: 536870912` (0x20000000) = linked ODBC table
- DSN names are case-sensitive in connection strings
- `SourceTableName` uses the format `schema.tablename` (e.g., `dbo.tMyTable`)
- `PrimaryKey` uses bracket notation; for a compound key, separate columns with `, ` (e.g., `[Id]` or `[Id], [NameID]`)

### Creating a new linked table or view link

Two things must happen before Access can resolve a new linked-table name:

1. **The object exists on the SQL Server side.** Deploy the table/view to the database the DSN points to.
2. **The tbldef file exists in `tbldefs/`** with the correct encoding (UTF-8 BOM + CRLF). Use `printf` rather than `Write`/`echo`, which strip BOM:
   ```bash
   printf '\xEF\xBB\xBF{\r\n  "Info": {\r\n    "Class": "clsDbTableDef",\r\n    "Description": "Linked Table"\r\n  },\r\n  "Items": {\r\n    "Name": "vwMyView",\r\n    "Connect": "ODBC;DSN=CRM;Trusted_Connection=Yes;APP=Microsoft Office;DATABASE=CRM;",\r\n    "SourceTableName": "dbo.vwMyView",\r\n    "Attributes": 536870912,\r\n    "PrimaryKey": "[Id]"\r\n  }\r\n}\r\n' > CRM.accdb.src/tbldefs/vwMyView.json
   ```

### How linked tables appear in `vcs-index.json`

Quirk: linked-table tbldef files are listed under the `"Tables"` component of `vcs-index.json` with an **`.xml`** extension regardless of the actual `.json` file on disk. For example, `tbldefs/vwMyView.json` is tracked as `"vwMyView.xml"` in the index. When grepping the index for a tbldef entry, search by base name only.

## Local Table Definitions (tbldefs/*.xml)

Local tables (stored inside the .accdb, not linked to SQL Server) are exported as XSD schema files:

```xml
<?xml version="1.0"?>
<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema"
            xmlns:od="urn:schemas-microsoft-com:officedata">
  <xsd:element name="dataroot">
    <xsd:complexType>
      <xsd:sequence>
        <xsd:element ref="tTableName" minOccurs="0" maxOccurs="unbounded"/>
      </xsd:sequence>
      <xsd:attribute name="generated" type="xsd:dateTime"/>
    </xsd:complexType>
  </xsd:element>
  <xsd:element name="tTableName">
    <xsd:annotation>
      <xsd:appinfo>
        <od:index index-name="PrimaryKey" index-key="ID " primary="yes"
                  unique="yes" clustered="no" order="asc"/>
        <od:tableProperty name="Orientation" type="2" value="0"/>
      </xsd:appinfo>
    </xsd:annotation>
    <xsd:complexType>
      <xsd:sequence>
        <xsd:element name="ID" minOccurs="1" od:jetType="autonumber"
                     od:sqlSType="int" od:autoUnique="yes" od:nonNullable="yes"
                     type="xsd:int"/>
        <xsd:element name="FieldName" minOccurs="0" od:jetType="text"
                     od:sqlSType="nvarchar" type="xsd:string"/>
      </xsd:sequence>
    </xsd:complexType>
  </xsd:element>
</xsd:schema>
```

- Common `od:jetType` values: `autonumber`, `text`, `longinteger`, `datetime`, `yesno`, `currency`, `double`, `memo`
- Local tables are typically used for temporary/working data (e.g., `tTEMP_*`, `tSelect_*`), lookup data, and system tables
- **Do not edit these unless necessary** — schema changes are usually made in SQL Server for linked tables

### Table Creation Rules

When creating new local tables, always follow these conventions:

1. **Every table MUST have an `ID` column as primary key** — `autonumber` type, non-nullable, auto-incrementing integer. This is the standard Access pattern for surrogate keys.

2. **Every foreign key column MUST have a corresponding relationship file** in the `relations/` directory. For example, if `tOrderItems` has a `OrderID` column referencing `tOrders.ID`, create a relationship file `relations/tOrderstOrderItems.json`.

3. **Foreign key columns** should be named `<ReferencedTable>ID` (without the `t` prefix) — e.g., `CustomerID` references `tCustomers.ID`, `OrderID` references `tOrders.ID`.

## Relationships (relations/*.json)

Table relationships define foreign key constraints, referential integrity, and join types between tables. Each relationship is a `.json` file in the `relations/` directory.

### File Naming Convention

The filename concatenates the parent table name and child table name: `<ParentTable><ChildTable>.json`

Example: `tOrderstOrderItems.json` (relationship from `tOrders` to `tOrderItems`)

For relationships with custom join types, append a suffix: `<ParentTable><ChildTable> - join type N.json`

### Relationship Format

```json
{
  "Info": {
    "Class": "clsDbRelation",
    "Description": "Database relationship"
  },
  "Items": {
    "Name": "tOrderstOrderItems",
    "Table": "tOrders",
    "ForeignTable": "tOrderItems",
    "Attributes": 2,
    "Fields": [
      {
        "Name": "ID",
        "ForeignName": "OrderID"
      }
    ]
  }
}
```

- `Name`: Relationship name (typically `<ParentTable><ChildTable>`)
- `Table`: The **parent** (primary key) table
- `ForeignTable`: The **child** (foreign key) table
- `Fields`: Array of field mappings. `Name` = field in parent table, `ForeignName` = field in child table
- `Attributes`: Controls referential integrity and join type (see below)

### Attributes Reference

The `Attributes` value controls referential integrity enforcement and join type:

| Attributes | Referential Integrity | Join Type | Description |
|---|---|---|---|
| `2` | No | Inner Join (default) | No enforcement — allows orphan records. Default for simple relationships without referential integrity. |
| `0` | Yes | Inner Join (default) | Enforced — prevents orphan records. Access rejects inserts/updates that violate the relationship. |
| `16777216` | Yes | Left Outer Join | Enforced. Include ALL records from the parent table (`Table`) and only matching records from the child table (`ForeignTable`). |
| `33554432` | Yes | Right Outer Join | Enforced. Include ALL records from the child table (`ForeignTable`) and only matching records from the parent table (`Table`). |

**When to use each:**
- `Attributes: 2` — Loose relationship, no data integrity checks. Use for optional or legacy associations.
- `Attributes: 0` — Standard enforced relationship. Use for required foreign keys (e.g., every order must have a valid customer).
- `Attributes: 16777216` — Left join. Use when the parent record may not have any child records but you still want to see it (e.g., show all customers, even those without orders).
- `Attributes: 33554432` — Right join. Use when the child record's parent may be missing but you still want to see it (e.g., show all order items, even if the order header is incomplete).

### Compound Key Relationships

For tables with composite foreign keys, add multiple entries to the `Fields` array:

```json
"Fields": [
  {"Name": "CompanyID", "ForeignName": "CompanyID"},
  {"Name": "BranchID", "ForeignName": "BranchID"}
]
```

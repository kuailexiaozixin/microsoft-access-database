# Queries Reference

## Query `.bas` Metadata File

Contains:
- `Operation` flag (1 = SELECT, etc.)
- `InputTables`, `OutputColumns`, `Joins` blocks for the visual designer
- Designer window coordinates at end of file (don't modify)
- Pass-through queries: `dbMemo "Connect"` and `dbMemo "SQL"` properties

### Single Joins Block Rule

**There must be at most one `Begin Joins` block per query `.bas` file.** When a query has multiple joins, list all join definitions inside a single `Begin Joins...End` block — do NOT use separate `Begin Joins` blocks for each join. A second `Begin Joins` block causes a fatal import error:

```
Error encountered at line N.
Expected: End of file.  Found: Joins.
```

Correct (multiple joins in one block):
```
Begin Joins
    LeftTable ="tOrders"
    RightTable ="tCustomers"
    Expression ="tOrders.CustomerID = tCustomers.ID"
    Flag =1
    LeftTable ="tOrders"
    RightTable ="tOrderItems"
    Expression ="tOrders.ID = tOrderItems.OrderID"
    Flag =1
End
```

Wrong (separate blocks — causes import error):
```
Begin Joins
    LeftTable ="tOrders"
    RightTable ="tCustomers"
    Expression ="tOrders.CustomerID = tCustomers.ID"
    Flag =1
End
Begin Joins
    LeftTable ="tOrders"
    RightTable ="tOrderItems"
    Expression ="tOrders.ID = tOrderItems.OrderID"
    Flag =1
End
```

## Query `.sql` File

Contains the actual SQL. For pass-through queries, this is the T-SQL sent to SQL Server.

Both files must stay synchronized. Editing `.sql` alone may not be sufficient if the `.bas` designer metadata contradicts it.

## `dbMemo "SQL"` and `SaveQuerySQL` Interaction

**When `SaveQuerySQL=true`** in `vcs-options.json` (the recommended setting), VCS reads query SQL exclusively from the paired `.sql` file during import. The `.bas` file must **NOT** contain a `dbMemo "SQL"` property. Including it causes a fatal import error:

```
Error encountered at line N.
Could not create or set the property SQL.
```

This happens because VCS sets the SQL property from the `.sql` file first, then the `.bas` file tries to set it again via `dbMemo "SQL"`, and Access rejects the duplicate assignment.

**Rules:**
- **`SaveQuerySQL=true`** (default/recommended): Never include `dbMemo "SQL"` in the `.bas` file. The `.sql` file is the single source of truth for the SQL text.
- **`SaveQuerySQL=false`**: The `.bas` file **must** contain `dbMemo "SQL"` because there is no `.sql` file.
- **Pass-through queries**: Always use `dbMemo "SQL"` in the `.bas` file (pass-through SQL is not split to `.sql` regardless of the setting). Also include `dbMemo "Connect"` for the ODBC connection string.

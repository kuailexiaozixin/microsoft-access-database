# Conditional Formatting (Forms & Reports)

> Cross-reference: [Forms Critical Rules](forms.md#critical-rules) and [Reports Critical Rules](reports.md#critical-rules) apply — especially Rule 3 (never edit binary data blocks).

`ConditionalFormat` and `ConditionalFormat14` are hex-encoded binary properties that appear on form and report controls (TextBox, ComboBox, etc.) when conditional formatting rules have been set via the Access designer. Example:

```
Begin TextBox
    ...
    ConditionalFormat = Begin
        0x0100000...
    End
    ConditionalFormat14 = Begin
        0x0100000...
    End
    ...
End
```

## Decoding ConditionalFormat Binary Blocks

The binary data inside `ConditionalFormat` and `ConditionalFormat14` blocks follows a structured format. When you encounter an existing block, attempt to decode it to understand the rules before removing it. `ConditionalFormat14` is the modern (Access 2007+) variant with the same logical structure but different header/version bytes.

### Header Structure

| Offset | Size | Field | Notes |
|--------|------|-------|-------|
| 0x00 | DWORD | Version | Format version identifier |
| 0x04 | DWORD | Total size | Total byte count of the binary block |
| 0x08 | DWORD | Condition count | Number of condition records that follow |

### Per-Condition Record Fields

Each condition record contains:

**Condition type** (byte):
| Value | Type |
|-------|------|
| `0x00` | Field Value |
| `0x01` | Expression |
| `0x02` | Field Has Focus |

**Operator** (byte, applicable when condition type = Field Value):
| Value | Operator |
|-------|----------|
| `0x00` | Between |
| `0x01` | Not Between |
| `0x02` | Equal To |
| `0x03` | Not Equal To |
| `0x04` | Greater Than |
| `0x05` | Less Than |
| `0x06` | Greater Than Or Equal |
| `0x07` | Less Than Or Equal |

**Formatting flags:** Bold, Italic, Underline (each as a flag/byte).

**BackColor** (background fill): COLORREF as little-endian DWORD (`0x00BBGGRR`).

**ForeColor** (font colour): COLORREF as little-endian DWORD (`0x00BBGGRR`).

**Expression text:** UTF-16LE encoded string (for Expression-type conditions or Field Value comparison values).

**Sort/order flag:** Possible ordering byte for rule evaluation precedence.

### "Compare to Other Records" (Data Bars)

Data bar conditions include additional fields:

| Field | Description |
|-------|-------------|
| Show Bar Only | Flag indicating whether to hide the value and show only the bar |
| Shortest Bar | Value representing the minimum bar length |
| Longest Bar | Value representing the maximum bar length |
| Type | Data bar type identifier |
| Value | Data bar value reference |
| Color | Bar fill colour as COLORREF |

## Rules: Decoding Allowed, Encoding Banned

**Decoding is allowed:** When encountering existing binary blocks, attempt to decode and document the rules found (condition type, operator, expressions, colours, formatting). If decoding fails or is ambiguous, ask the user to clarify.

**Encoding/copying is still banned:** Never manually construct, modify, or copy binary `ConditionalFormat` or `ConditionalFormat14` hex blocks. All conditional formatting must be applied programmatically using VBA via the `FormatConditions` collection.

## Workflow: Modifying Conditional Formatting on a Control

1. **Decode** the existing binary block to extract the current rules (condition types, operators, expressions, colours, formatting flags). If decoding fails, ask the user to provide the rules.
2. **Document** the extracted rules (conditions, colours, formatting) so nothing is lost.
3. **Delete** the binary `ConditionalFormat = Begin...End` and `ConditionalFormat14 = Begin...End` blocks from the `.bas` file entirely.
4. **Recreate** using VBA `FormatConditions` collection in the `.cls` code-behind (see patterns below).
5. **Then** make any additional modifications the user requested (adding new rules, changing colours, etc.).

## Workflow: Copying Conditional Formatting to Another Control

1. **Decode** the source control's binary block to extract the rules (or ask the user if decoding fails).
2. **Delete** the source binary block (optional — if converting the source control to VBA too).
3. **Create** the VBA `FormatConditions` on the target control using the decoded rules in the `.cls` code-behind.

## VBA Patterns for Conditional Formatting

### Choosing the Right Event

| Form/Report Type | Event | Why |
|---|---|---|
| **Single-record form** | `Form_Current` | Fires each time a record gets focus |
| **Continuous/Datasheet form (list form)** | `Form_Load` | Rules are set once; Access auto-evaluates them across all visible rows. `Form_Load` fires after the recordset is loaded but before display. |
| **Report** | `Report_Open` | Rules are set once before the report renders |

**Important for list forms:** `FormatConditions` are rule-based, not row-by-row — Access re-evaluates them automatically for every visible row. You only need to set them up once. However, certain actions **destroy** format conditions and require re-applying them:
- `Me.Requery` or `Me.RecordSource = ...` (rebuilds the form's recordset)
- `DoCmd.ApplyFilter` in some contexts
- Any operation that effectively recreates the controls

For list forms that requery or change RecordSource, call `ApplyConditionalFormats` after those operations as well.

### List Form Pattern (Continuous/Datasheet)

```vba
Private Sub Form_Load()
    ApplyConditionalFormats
End Sub

Private Sub ApplyConditionalFormats()
    ' Clear any existing format conditions
    Me.txtFieldName.FormatConditions.Delete

    ' Add conditions (up to 50 in Access 2007+)
    ' Type 0 = acFieldValue, Type 1 = acExpression

    ' Example: Field value > 100 — red background
    With Me.txtFieldName.FormatConditions.Add(acFieldValue, acGreaterThan, 100)
        .BackColor = RGB(255, 200, 200)
        .ForeColor = RGB(139, 0, 0)
    End With

    ' Example: Expression-based — bold green when above target
    With Me.txtFieldName.FormatConditions.Add(acExpression, , "[Amount] > [Target]")
        .BackColor = RGB(200, 255, 200)
        .FontBold = True
    End With
End Sub

' If the form requeries, re-apply after the requery
Private Sub btnRefresh_Click()
    Me.Requery
    ApplyConditionalFormats
End Sub
```

### Single-Record Form Pattern

```vba
Private Sub Form_Current()
    ApplyConditionalFormats
End Sub

Private Sub ApplyConditionalFormats()
    Me.txtStatus.FormatConditions.Delete

    With Me.txtStatus.FormatConditions.Add(acFieldValue, acEqual, "Overdue")
        .BackColor = RGB(255, 200, 200)
        .ForeColor = RGB(139, 0, 0)
        .FontBold = True
    End With
End Sub
```

### Report Pattern

```vba
Private Sub Report_Open(Cancel As Integer)
    ApplyConditionalFormats
End Sub

Private Sub ApplyConditionalFormats()
    Me.txtTotal.FormatConditions.Delete

    With Me.txtTotal.FormatConditions.Add(acFieldValue, acGreaterThan, 10000)
        .BackColor = RGB(255, 255, 200)
        .FontBold = True
    End With
End Sub
```

## Why This Matters

- Hex blobs are **not diffable** — changes are invisible in code review
- They **cannot be merged** — any conflict requires discarding one side entirely
- They **cannot be hand-edited** — a single byte error corrupts the control
- VBA-based formatting is **readable, diffable, and mergeable**

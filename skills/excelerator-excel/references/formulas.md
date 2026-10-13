# Formula and SQL reference

## Worksheet functions

| Formula | Result and defaults |
|---|---|
| `SFVALUE(sql, [refresh=0])` | One column, at most one row. |
| `SFQUERY(sql, [headers=TRUE], [refresh=0], [maxRows=10000])` | Table with rows first, columns second, and optional headers. |
| `SFDATABASES([refresh=0])` | Database names with a header. |
| `SFSCHEMAS(database, [refresh=0])` | Schema names with a header. |
| `SFTABLES(database, schema, [refresh=0])` | Table/view names and types. |
| `SFCOLUMNS(database, schema, table, [refresh=0])` | Column name, type, numeric precision, scale, and nullability in source order. |
| `SFSTATUS([queryCell])` | An 11-row, 2-column status table. `queryCell` is one formula anchor in this workbook, not the entire spill. |

Inputs are values or single-cell references. `headers` accepts TRUE/FALSE or 1/0. `maxRows` must be a whole number from 1 to 1,048,575; the separate 250,000-cell result limit still applies. Each formula anchor has its own cache. Metadata formulas use the same refresh behavior and return headers.

Status labels are **State**, **Message**, **Last success**, **Rows**, **Columns**, **Server**, **Database**, **Schema**, **Role**, **Warehouse**, and **User**. Rows excludes headers; time uses the computer's local time. A standalone `SFSTATUS()` may need worksheet recalculation to reflect a connection change. A status formula referencing an output recalculates as its dependent.

## Example discovery and report layout

Use these formulas only in inspected, empty locations; names below are examples to replace with discovered objects:

```excel
=SFDATABASES($B$7)
=SFSCHEMAS("FINANCE",$B$7)
=SFTABLES("FINANCE","PUBLIC",$B$7)
=SFCOLUMNS("FINANCE","PUBLIC","GL_ENTRIES",$B$7)
```

One recurring schedule can use entity in `B4`, period start in `B5`, refresh token in `B7`, row limit in `B8`, scalar SQL in `B11`, and detail SQL in `B12`:

```excel
=SFVALUE(B11,$B$7)
=SFQUERY(B12,TRUE,$B$7,$B$8)
=SFSTATUS(A18)
```

The last formula assumes the table formula is at `A18`. Its status area needs 11 empty rows and 2 empty columns. Place the scalar formula separately. Increasing `B7` refreshes both data outputs; it is not a timer. Formula changes can fetch data immediately when automatic calculation is enabled.

## SELECT templates

Adapt identifiers and filters using discovered metadata. This example returns monthly account totals for one entity; it assumes `posting_date` represents the intended accounting date and that the amount/sign convention is already known:

```sql
SELECT
    CAST("ACCOUNT_CODE" AS VARCHAR) AS "account_code",
    "ACCOUNT_NAME" AS "account_name",
    SUM("AMOUNT") AS "amount"
FROM "FINANCE"."PUBLIC"."GL_ENTRIES"
WHERE "ENTITY" = 'EXAMPLE_ENTITY'
  AND "POSTING_DATE" >= TO_DATE('2026-09-01')
  AND "POSTING_DATE" < TO_DATE('2026-10-01')
GROUP BY "ACCOUNT_CODE", "ACCOUNT_NAME"
ORDER BY "ACCOUNT_CODE"
```

For a scalar total, select only `SUM("AMOUNT")` using the same source and filters. An empty aggregate may be SQL NULL, which becomes `""`; show that distinction instead of inventing a zero. For timestamps, choose boundaries in the report's intended timezone before building SQL.

Quote each identifier component with double quotes; double embedded double quotes. Metadata function arguments instead take the raw exact name, for example `FINANCE`, not `"FINANCE"` as part of the cell value. Quoted SQL identifiers are case-sensitive; do not uppercase names that were discovered with mixed case.

For an entity input in `B4`, this Excel expression produces escaped SQL literal content, without outer apostrophes:

```excel
=SUBSTITUTE(SUBSTITUTE(B4,"\","\\"),"'","''")
```

Wrap that content in apostrophes when assembling SQL. Format date inputs with `TEXT(B5,"yyyy-mm-dd")` and use `TO_DATE`. Do not concatenate user input into identifier positions unless it is validated against the discovered object and quoted as an identifier.

## Exact values and numeric calculations

A Snowflake `NUMBER(38,2)` column returns text because its declared precision exceeds Excel's 15 significant digits, even when a specific amount is small. Exact text is suitable for display and comparison, but worksheet `SUM` will not treat it as a normal numeric range. Aggregate in SQL when possible.

When the report explicitly needs numeric cells and the range and scale are appropriate, project `CAST(SUM("AMOUNT") AS NUMBER(15,2))`. That cast can round or fail on overflow; make the conversion visible in the report's SQL and explanation. To preserve an exact source value when ODBC floating-point conversion is a problem, project `CAST("AMOUNT" AS VARCHAR)`. Do not convert precise text through Double and claim the original value survived.

Dates are returned as date values; format their output cells without changing the query's meaning. Text cells are limited to 32,767 characters; shorten larger values deliberately in SQL if the report permits it.

## Error recovery

| Observed result | Action |
|---|---|
| `#NAME?` | Check that Excelerator 1.2+ is loaded in this Excel instance and that the tool evaluates VBA functions. If installation is missing, explain the requirement; do not change macro security settings as a query workaround. |
| `#N/A` | Read status. Have the accountant connect/reconnect in this workbook, then change the refresh token. |
| `#VALUE!` | Read status for SQL syntax, invalid inputs, scalar shape, precision, unsupported values, or result-limit errors. Fix the identified cause rather than changing unrelated settings. |
| `#SPILL!` | Inspect occupied/merged cells and Excel Tables. Move the formula to a suitable empty area or resolve an obstruction within the user's authorized edits; preserve existing work. |
| Success status with a spill error | The database query succeeded, but Excel cannot display its array. Resolve the spill separately. |
| Old last-success time with Error state | The latest attempt failed. Do not present the historical timestamp or previously saved values as a successful refresh. |
| Refresh needed | Recalculate or change refresh after the connection/context change; reconnect if connection status requires it. |
| Not evaluated | Confirm the referenced cell is the intended query formula anchor and recalculate it. |

Saved workbook values and session caches are distinct: closing Excel clears the in-memory caches and connection, while a file may still show old saved values. Evaluate and inspect current status before reporting freshness.

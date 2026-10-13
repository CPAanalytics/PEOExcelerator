# Worksheet formulas for schedules and Excel agents

Install one of the version 1.2 add-ins in Microsoft 365 Excel for Windows. Click **Connect** in the workbook that will contain the formulas. Authentication remains interactive; a worksheet formula never opens a login dialog or asks an agent to handle credentials.

## Formula reference

| Function | Arguments | Result |
|---|---|---|
| `SFVALUE` | `sql, [refresh=0]` | One value from one column and at most one row. |
| `SFQUERY` | `sql, [headers=TRUE], [refresh=0], [maxRows=10000]` | A table with rows first and columns second. |
| `SFDATABASES` | `[refresh=0]` | Accessible database names, with a header. |
| `SFSCHEMAS` | `database, [refresh=0]` | Accessible schema names, with a header. |
| `SFTABLES` | `database, schema, [refresh=0]` | Table/view names and object types. |
| `SFCOLUMNS` | `database, schema, table, [refresh=0]` | Column names, data types, numeric precision/scale and nullability in source order. |
| `SFSTATUS` | `[queryCell]` | Labels and values describing the connection or a query's latest attempt. |

Supply exact object names without surrounding SQL quotes to discovery formulas. The add-in escapes identifiers and literals. Metadata is limited to objects visible to the connected Snowflake role. Data functions accept one `SELECT` or `WITH … SELECT`, including an optional trailing semicolon. They reject additional statements and top-level write commands. This statement guard does not replace Snowflake permissions: use the role appropriate to your reporting access, especially for functions that may have side effects.

SQL and other inputs can be literal values or single-cell references. Error-valued inputs, multi-cell argument ranges and nonintegral `maxRows` are rejected. `maxRows` must be between 1 and 1,048,575; a separate 250,000-cell limit applies to each result, including headers. Overflow returns an error, never a partial accounting schedule. These limits bound the returned data, not Snowflake scan work; filter or aggregate queries to reduce database work.

SQL NULL and empty scalar results return `""`; numeric zero stays zero. An empty table returns only its headers, or `""` if headers are disabled. More than one scalar row or column is an error. Dates are returned as date values; format the output cells as needed. Formula results do not change worksheet formatting. Strings longer than Excel's 32,767-character cell limit are rejected.

## Refresh and connection behavior

Use a shared refresh cell such as `B7`:

```excel
=SFVALUE(B11,$B$7)
=SFQUERY(B12,TRUE,$B$7,10000)
=SFSTATUS(A18)
```

The first evaluation fetches data. Changing SQL, result options or the refresh value fetches again. Ordinary recalculation reuses the result at each formula anchor, including an error from a failed attempt. Separate anchors have separate caches. Increase `B7` by 1 to retry a failed query or refresh the schedule. The functions are nonvolatile and do not refresh on a timer.

Caches live in memory, are cleared on connection/session changes, and are released when workbooks close. After restarting Excel, connect again before refreshing. Saved workbook values are not proof of current data; inspect status after recalculation. `SFSTATUS` itself makes no database call and reads no passwords. It reports the latest attempt, last successful refresh time in the computer's local time, data-row count excluding headers, column count, and effective session context. On failure, it keeps the historical last-success timestamp while marking the current attempt as an error.

The underlying add-in has one shared interactive connection. The new formulas associate that session with the workbook connected through **Connect**. Connecting another workbook invalidates previous caches; switch back and connect again to query the first workbook. The functions resolve context from the formula's workbook, regardless of the active sheet. Edited connection settings require reconnecting; formulas never silently use a session belonging to another workbook.

Status that references a query anchor recalculates as a dependent formula. A standalone `SFSTATUS()` may need worksheet recalculation to reflect a connection change. Changing the role, warehouse, or default schema through the existing commands invalidates cached query results. Use fully qualified table names to make report SQL clear.

## Accounting data and precision

Text account numbers retain leading zeros. A number that was already stored numerically in Snowflake cannot recover lost leading zeros through formatting alone.

The add-in does not round amounts or apply currency formats. Decimal columns whose declared precision exceeds 15 return as text to avoid Excel's numeric precision limit—even when a particular value is small. Large 64-bit integers also return as text. Text amounts do not participate in `SUM` the same way numeric cells do.

If a numeric worksheet result is required, choose an explicit SQL projection whose precision and scale fit the report, for example `CAST(SUM(amount) AS NUMBER(15,2))`. This is an explicit reporting choice: the cast can round fractional digits or reject an out-of-range amount. Keep exact text where that conversion would be inappropriate. If the ODBC driver has already converted a high-precision value to floating point, the formula returns an error; use `CAST(column AS VARCHAR)` to retrieve it exactly. Text representations follow the local VBA number format when the driver supplies a Decimal value.

## Agent workflow

An agent needs permission and tooling to edit worksheet cells and formulas in the running Excel application. Provider integrations differ; support must be verified in the actual Claude or ChatGPT environment. These functions do not add an MCP server, chat panel, or provider-specific integration.

1. Insert `=SFSTATUS()` in an empty area. If it says **Not connected**, ask the accountant to use **Connect** in that workbook. Do not request, read, or write credentials.
2. Discover the source using `SFDATABASES`, `SFSCHEMAS`, `SFTABLES` and `SFCOLUMNS`, placing each spill in a clear area. Use exact discovered names and inspect types before choosing numeric conversions.
3. Put a single SELECT in a dedicated SQL cell. Prefer explicit columns, fully qualified names, an entity filter, and an unambiguous date range.
4. Insert `SFVALUE` for a single metric or `SFQUERY` for a detail table. Reference a shared refresh cell. Keep the spill area clear of notes, formulas, merged cells and Excel Tables.
5. Insert `SFSTATUS(queryAnchor)` beside the output. Inspect both Excel's formula error and the query status before describing the data as current or complete.
6. Increment the refresh cell when fresh data is needed. Ordinary recalculation and an unchanged token intentionally reuse the latest result.

For text values interpolated into SQL, escape backslashes and apostrophes. The example workbook uses `SUBSTITUTE(SUBSTITUTE(entityCell,"\","\\"),"'","''")`. Build dates with `TEXT(dateCell,"yyyy-mm-dd")` and `TO_DATE`. This release accepts SQL text rather than a separate bind-parameter API.

## Troubleshooting

| Symptom | Action |
|---|---|
| `#NAME?` | Install/load a rebuilt version 1.2 add-in with macros enabled. The original packages do not contain these functions. |
| `#N/A` | Inspect status; connect this workbook again for missing, closed or expired sessions, then change the refresh cell. |
| `#VALUE!` | Inspect status for SQL, input, shape, precision or result-size errors. Correct the query or inputs, then refresh. |
| `#SPILL!` | Clear the intended spill area or move the formula. The add-in never overwrites occupied cells. |
| Query status says Success but output has a spill error | Status describes database execution; Excel separately controls whether the array can spill. |
| Amounts appear as text | Inspect source precision. Choose an explicit numeric cast only when suitable for the report. |
| Connecting one workbook stops another workbook's formulas | Connect the workbook you want to query; the existing shared session is associated with one workbook at a time. |

The old **Query** ribbon command still clears its results worksheet. Upload and rollback workflows remain unchanged; the new formulas are intended for reading data.

## Validation and release checks

`scripts/Build-Addins.ps1` compiles the entire release VBA project before testing and again after removing test components. Obsolete developer-only `TestHarness` and `Testing` modules remain in source but are excluded from releases. It runs the actual VBA in Excel using a fake connection and disconnected ADO recordsets. Tests cover scalar shape, empty/NULL/zero results, orientation, headers, exact result limits, precision, cache behavior, retries, connection changes, multiple workbooks, worksheet spills, and recalculation of visible/hidden worksheets. Tests require VBA project access, and the script does not change that setting. It preserves startup/close events as disabled during its isolated build and does not run live upload tests.

Run `scripts/Test-Packages.ps1` after building to verify startup, both read-only modes, formula resolution, disconnected-session errors, and standalone/query status in the saved packages. It performs no Snowflake queries.

Before deploying in a connected environment:

- Connect with a Snowflake reporting role; verify discovery, a scalar aggregate, a filtered detail query, and the example schedule adapted to the real schema.
- Check NUMBER, large integer, text identifier and date handling with the installed ODBC driver. Verify exact values against Snowflake.
- Confirm ordinary recalculation does not issue new queries and changing refresh does. Test an expired session and a result exceeding the limits.
- Exercise the existing connection, ribbon query and upload/rollback entrypoints in an appropriate test environment.
- Separately have Claude and ChatGPT insert and evaluate discovery, scalar, table and status formulas, then change the refresh cell. Record the provider environment and outcome before claiming compatibility.

Local automated checks do not establish live Snowflake or provider compatibility.

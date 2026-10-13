---
name: excelerator-excel
description: Use the Excelerator add-in to discover Snowflake schemas, retrieve scalar or table results, and refresh accounting schedules through worksheet formulas in Microsoft 365 Excel for Windows. Use for work in an existing Excel workbook with live cell and formula access. Does not cover database write-back or add-in development.
---

# Excelerator in Excel

Use Excelerator 1.2 or later through its worksheet functions. Build the requested report in the user's workbook while preserving surrounding formulas, formatting, notes, and existing schedules. The formulas read Snowflake; authentication remains with the accountant.

## Establish access and context

- Use the user's existing Excel session and the intended workbook. Confirm the available Excel tool can write formulas, recalculate them in desktop Excel, and read evaluated values and spill ranges. A file-editing library or a service that only stores formulas cannot evaluate this VBA add-in.
- If using local automation, attach to the intended running Excel instance. A new COM automation instance has a separate authentication session and may skip startup add-ins. Do not assume it has the user's installed functions or Snowflake connection.
- Inspect the workbook layout before choosing output locations. `SFSTATUS` needs an empty 11-row by 2-column area. Discovery and query functions also spill; use an empty area outside Excel Tables and merged cells. Reference the formula anchor when inspecting status.
- Insert `=SFSTATUS()` and recalculate. If **Not connected**, have the accountant select **Snowflake → Connect** in this workbook. Do not request or read passwords, SSO tokens, or hidden credential cells. Resume once status confirms the intended server, database, schema, role, and warehouse.
- A shared session belongs to one connected workbook at a time. Connecting a different workbook invalidates previous query caches. Reconnect the intended workbook if its settings or connection context changed.

If the tool cannot evaluate these formulas, explain the missing capability and provide the SQL/formulas as a draft. Do not describe cached file values or unexecuted formulas as newly queried data.

## Discover and retrieve data

Use the function signatures and SQL examples in [references/formulas.md](references/formulas.md) when constructing formulas or diagnosing errors.

1. Discover the actual accessible objects with `SFDATABASES`, `SFSCHEMAS`, `SFTABLES`, and `SFCOLUMNS`. Supply exact object names without surrounding SQL quotes. Inspect types and numeric precision before selecting report columns; example schema names are not evidence of real tables.
2. Establish the report's entity, period, grain, and amount/sign convention from the request and existing workbook. Ask only about material ambiguity that those sources cannot resolve. Use fully qualified identifiers and explicit columns; quote each SQL identifier separately and escape embedded double quotes.
3. Keep SQL in a dedicated cell and report inputs in identifiable cells. Use a single `SELECT` or `WITH ... SELECT`; escape interpolated text and use unambiguous date bounds. These functions do not have a bind-parameter argument.
4. Use `SFVALUE(sqlCell, refreshCell)` for one metric, or `SFQUERY(sqlCell, TRUE, refreshCell, maxRows)` for a table. Use a shared numeric refresh cell for related outputs. Place results where their spill can grow without replacing existing work.
5. Put `SFSTATUS(queryAnchor)` beside each output that needs freshness information. Recalculate and read both the evaluated result and status before using the data in the report.

Do not use the legacy **Query** ribbon workflow to populate a schedule beside existing work: it clears its results worksheet. Do not use upload, rollback, arbitrary VBA query helpers, or database write commands as a workaround for a failed read query. A request for write-back needs a separately scoped workflow.

## Refresh and interpret results

- The first evaluation, changed SQL/options, or changed refresh value fetches data. Ordinary recalculation reuses that anchor's in-memory result, including failures. To fetch again, increment the existing refresh cell once; do not poll by repeatedly recalculating or changing it.
- On a query error, inspect status, address its concrete cause, then retry using changed inputs or the refresh value. If the same failure persists, report it and the needed action. Connection failures require the accountant to reconnect; do not loop login or query attempts.
- Status records the latest attempt and the historical last-success time. An error with an old timestamp is a failed current refresh. **Success** describes database execution; still inspect the output for Excel spill or display errors.
- Results default to 10,000 data rows and are limited to 250,000 cells including headers. Oversized results fail completely. Filter or aggregate to the requested report grain; do not add a limiting clause merely to conceal missing report rows.
- `SFVALUE` rejects multiple rows or columns. SQL NULL and empty scalar results become `""`; do not automatically treat them as zero. An empty table returns headers only, or `""` when headers are disabled.
- Preserve text identifiers and leading zeros. Decimal columns with declared precision above 15 and large integers can return exact text. Do not silently coerce amounts, use `SUM` on text as if it were numeric, or introduce rounding to remove an error. Use an explicit numeric SQL cast only when its precision and scale fit the report requirements. If the driver already returned an imprecise floating-point value, query an explicit `VARCHAR` cast instead.

## Complete the handoff

Leave the SQL cells, report inputs, refresh control, result anchors, and status areas identifiable. Tell the accountant which cells to edit and how to refresh. Report the effective source/context, entity and period, returned row count, latest attempt, last successful refresh, and any missing data or numeric conversions relevant to the task. Confirm surrounding workbook content is preserved. Claim live Snowflake/provider verification only when the agent actually evaluated and inspected the results in that environment.

# Excelerator: Use Snowflake in Excel

Excelerator is an Excel Add-In to help you pull data from Snowflake into Excel and push new or updated data from Excel into Snowflake.
Excelerator is only compatible with the Windows operating system, not the MacOS.
The worksheet functions introduced in version 1.2 require Microsoft 365 Excel for Windows with dynamic arrays.
View a [demo of Excelerator](https://youtu.be/csmS5V2ONr8) and see it [upload over 1 million rows (77MB)](https://www.youtube.com/watch?v=MLDnWLd1EHA). 

Example use-cases include:
* financial budgeting write-back,
* look-up table maintenance,
* pricing analysis,
* and more!

| :exclamation:        | Excelerator is not a supported product by Snowflake or any company. Excelerator will write data to the Snowflake database and should be used with great care. Use at your own risk.     |
|---------------|:------------------------|

## Get Started with Excelerator
Important: When following the instructions below, make sure to 'Unblock' the Addin after downloading. The details are in the link located in Step 2 below.

To get started, you'll need to:

* Install the Windows 64-bit or 32-bit ODBC driver for Snowflake
* Install the Add-In in Excel (below)
* [Connect your Snowflake account](#use-excelerator)
* [Query Snowflake](#execute-a-query) and optionally [write data back](#write-data-to-snowflake)

## Install the ODBC Driver
The Excel Add-In requires the ODBC driver. To download the driver, go to: https://sfc-repo.snowflakecomputing.com/odbc/index.html
Select the version of the ODBC driver, 64-bit or 32-bit, based on the local Excel installation. Make sure it matches or the Add-In will not work.
Once downloaded, install the driver. 

## Install Excelerator

### Step 1: Set Required Privileges

Below are the privileges required for each capability.

To **query** data:
| **Object** | **Privilege**        |
|------------|----------------------|
| Database   | USAGE                |
| Schema     | USAGE|
| Table      | SELECT               |

\*Stage can be provided in the login instead

To **upload** data requires everything in query, plus:

| **Object** | **Privilege**            |
|------------|--------------------------|
| Schema     | CREATE TABLE,CREATE STAGE|
| Table      | INSERT, UPDATE, TRUNCATE |

To **rollback** data requires the schema privileges from upload, plus:

| **Object** | **Privilege** |
|------------|---------------|
| Table      | Ownership     |
  

Optional
The following script will create a new role with the proper privileges, except the table level privileges: `SnowflakeExcelAdd-In_Create_Role.sql` 
You’ll have to update the script before executing it with the information specific to your environment. In the script, you will be providing the role of the user that will be using this Add-In. The script will assign the new to the existing role, which will inherit all the privileges defined in the script.

### Step 2 – Install Excel Add-in

There are 2 versions of the Excel Add-in, one for only reading data from Snowflake and one for reading and writing data to snowflake. The Excel Add-ins are Excel files with an extension of ".xlam". The .xlam file for the read-only version is called "SnowflakeExcelAdd-InReadOnly.xlam". The full read-write version is called "SnowflakeExcelAdd-In.xlam". These files are both stored in the repository. In order to install these add-ins, follow the instructions here: https://exceloffthegrid.com/install-uninstall-excel-add/. 


The Excelerator is now available on the Home tab of the Ribbon.

## Use Excelerator

With Excelerator installed, now you need to connect it to Snowflake. 

![](images/image4.jpg)

First, confirm you can find the appropriate buttons within the "Home" tab.

### Connection Parameters

Click the `connect` button and enter your Snowflake connection information.  There are two security types standard: login/password and SSO.  We recommend using SSO for more secure authentication.

![](images/image5.png)

The user's default role and warehouse will be used. If the user does not have a default warehouse, they will be prompted to enter one.
Both the role and warehouse can be changed by clicking on the 'Config' menu item in the Excel ribbon.

### Overview

Now we'll walk you through each section of the Add-In. You can work with data sourced from other data sources within Excel (such as files) or you can pull data from Snowflake into Excel.  You can manipulate the data using VBA scripts and vlookups, but be sure to write that final data back to Snowflake! The Add-In handles both reading and writing scenarios.

### Execute a Query
Pull data into Excel by selecting the “Query" button. 
A search dialog will open that allows you to select a database, schema and table in order to download data. Once you select a table you can then choose which columns you would like to reutrn from the selected table. You can select all the columns by clicking the "All" button, or select a subset of columns by clicking the "Choose" button. Once selected, a SQL statement is created and entered in the bottom text area. This SQL statement can be manually updated.
Click the "Execute" button to execute the query in Snowflake and pull down the result set.

![](images/image6.jpg)

You'll get results that look similar to this:

![](images/image7.png)

Repeat as-needed to gather data from Snowflake into your Excel sheet.

### Query from worksheet formulas

Use formulas to build recurring schedules without clearing surrounding cells. Install either rebuilt add-in, open your workbook, and click **Connect** once. An Excel agent can then discover tables, write SQL into a cell, insert formulas, and change a refresh cell without operating dialogs.

```excel
=SFVALUE("SELECT COUNT(*) FROM FINANCE.PUBLIC.GL_ENTRIES",$B$7)
=SFQUERY(B12,TRUE,$B$7,10000)
=SFSTATUS(A18)
```

`SFVALUE` returns one column and at most one row. `SFQUERY` spills rows and columns from its formula cell, with headers by default. Changing SQL or formula inputs fetches new data. Changing the optional refresh value also fetches again; ordinary recalculation reuses results, including failed attempts. Increment a shared refresh cell to refresh a schedule.

Start with [the recurring schedule workbook](examples/RecurringSchedule.xlsx). Its example table and columns must be adapted to your Snowflake schema. See [the formula and agent guide](docs/Worksheet-functions.md) for discovery formulas, precision handling, troubleshooting, and validation instructions.

The existing **Query** ribbon workflow still replaces the contents of its configured results sheet. Use the new worksheet formulas when working alongside existing schedules, formulas, or notes.

### Hand off the Excelerator skill

The reusable [Excelerator skill](skills/excelerator-excel/SKILL.md) guides an Excel agent through connection checks, schema discovery, scalar/table queries, controlled refresh, and accounting precision. Hand off the entire `skills/excelerator-excel` folder, including its `references` directory. It is self-contained and does not require the recipient to have this repository's developer tools.

Example request: “Read `excelerator-excel/SKILL.md` and use it to build a monthly account schedule in this workbook, preserving my existing formulas and notes.” The agent needs tooling that can edit and evaluate formulas in the running Windows Excel instance; supplying the skill alone does not create that connection. Database write-back is outside this skill's scope.

### Build and verify the add-ins

Run these commands from PowerShell in the repository. Excel must allow **Trust access to the VBA project object model**; the scripts do not change that setting. Close any installed copies of these add-ins before replacing them.

```powershell
./scripts/Build-Addins.ps1 -TestOnly
./scripts/Build-Addins.ps1
./scripts/Test-Packages.ps1
./scripts/Build-Example.ps1
```

The build imports the release source into a temporary copy, compiles the entire release VBA project, runs offline Excel/ADO integration tests, removes test components, compiles again, and verifies both packages before replacing the repository's `.xlam` files. The obsolete developer-only `TestHarness` and `Testing` modules remain in source but are excluded from release packages. The build does not connect to Snowflake or run live upload tests. To place packages elsewhere, pass `-OutputDirectory`.

With Node.js installed, `node tests/run-offline.cjs` checks the production SQL scanner and result conversion using Windows Script Host and disconnected ADO recordsets. Its narrow VBA-to-VBScript adaptations do not replace the Excel integration tests.

### Write Data to Snowflake

To write data to Snowflake, click the "Upload" button in the ribbon.

![](images/image8.png)

#### Select Upload Table
In the top section, choose the database, schema and table to upload the data to.

##### Basic Upload section
There are 3 basic ways to upload data into a table:
* Update existing rows & Append new rows 
* Append data - This will insert all the data into the table.
* Replace all data in the table with the data from the spreadsheet - All the data will be deleted from the table and then the new data will be inserted.

When selecting the first option, you will have to define the columns that represent the table key. The key defines how each row is unique, and is used to update existing rows. This can be done by entering the column position, for example: A,B,C.  In this case the first 3 columns in the excel sheet and target table will be the key.

With any of the above options, if a column exists in the spreadsheet, but does not exist in the table, it will be added automatically to the Snowflake table. The data type can be specified by clicking the "Define Data Types" button in the ribbon. If you prefer to have Snowflake determine the data types, select the "Auto-generate data types" in the "Advance Options" section below.

#### "Advanced Options" section 
This section allows more options:
* Create a new table - This will create a new table with the columns specified in the first row.
* Recreate and existing table - This will drop the table and recreate the table with the columns specified in the first row.

### Specifying Data Types 
For any new columns, the data type needs to be manually specified by clicking the "Define Data Types" button in the ribbon. A new row will be created in the first line of the spreadsheet containing a drop down of the Snowflake supported data types. Choose the appropriate data types for each column.

![](images/image9.jpg)

| :wrench: Limitations|
|:---------------------------|
| Timestamps don’t handle anything less than seconds. To get around this, manually cast the time to a varchar in the SQL. |



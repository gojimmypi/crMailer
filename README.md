# crMailer

Command-line tool for PDF rendering of a Crystal Report file and sending to recipient list.

Reports are typically created in a development environment. Over the course of time, production server names also change.
This library sets the datasource login (or more specifically the `TableLogOnInfo`) for a given report, and all embedded subreports, 
to a new server at runtime. (database names are asssumed to remain the same) 

Reports that span different servers or datasources will need to be adapted, 
or have the `SetCrystalDocumentLogon` code disabled to not change datasources. 

Uses [Microsoft SQL Server](https://docs.microsoft.com/en-us/sql/sql-server/?view=sql-server-ver15) `msdb.dbo.sp_send_dbmail` 
to [send email](https://docs.microsoft.com/en-us/sql/relational-databases/system-stored-procedures/sp-send-dbmail-transact-sql?view=sql-server-ver15), 
but could be easily adapted to use any SMTP relay. This `SQLHelper` library is only used for the call to the mailer.

## Requirements

This is a [Microsoft Visual Studio](https://visualstudio.microsoft.com/vs/) 2019 Project.

The [SAP Crystal Reports Runtime Redistributable Package](https://help.sap.com/viewer/0d6684e153174710b8b2eb114bb7f843/SP21/en-US/45b285716e041014910aba7db0e91070.html) 
needs to be installed locally, specifically these DLLs:

```
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using CrystalDecisions.ReportAppServer.CommLayer; // not used directly, but this is needed in Project References.
```

The [PatternsAndPractices Microsoft.ApplicationBlocks.Data](https://github.com/gojimmypi/PatternsAndPractices/tree/master/Microsoft.ApplicationBlocks.Data) is used for SQL Access.
See also the [Stack Exchange Dapper](https://github.com/StackExchange/Dapper) that works extremely well with [.NET Core](https://docs.microsoft.com/en-us/dotnet/core/) apps. 
Undetermined if it will work here.  The Crystal Reports runtime does not currently work with DotNetCore.

This app was developed with the Crystal Reports Developer Edition installed locally, with the runtime installed at the server.

To use this code, assuming there's a `c:\workspace` directory:

```
c:
cd\workspace\
git clone https://github.com/gojimmypi/crMailer.git
mkdir -p bin64
:: copy precompiled binaries such as CR runtime and Microsoft.ApplicationBlocks.Data into .\bin64
```

Set the desired parameters in the [App.config](./App.config) file.

Reminder: The file attach feature in `SendReport()` requires the SQL Server Service account to have file-level access for `OutputFileWithPath` at the local SQL server.
("C:" is the local drive at the server where `sp_send_dbmail1` runs, not necessarily where the code is running!)

Best sure `Save data with report` is *unchecked* in the Crystal Report file. See [#2](https://github.com/gojimmypi/crMailer/issues/2)

See also the [Crystal Reports Downloads](https://www.crystalreports.com/download/), [Free Viewer](https://www.sap.com/cmp/td/sap-crystal-reports-viewer-trial.html) 
and [SAP Crystal Reports, version for Visual Studio](https://www.sap.com/cmp/td/sap-crystal-reports-visual-studio-trial.html).


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Configuration;
using System.Threading;
// using CrystalDecisions.ReportAppServer.CommLayer; // not used directly, but this is needed in Project References.
//
// be sure to set "copy local" = true in the Project:
// see https://stackoverflow.com/questions/38025601/could-not-load-file-or-assembly-crystaldecisions-reportappserver-commlayer-ver

namespace crMailer
{
    class Program
    {
        static string crReportName = "(report name not specified)";
        const string REPORT_SUBDIRECTORY = "reports";
        static string OutputFileWithPath;
        static ReportDocument crReportDocument;
        static ConnectionInfo crConnectionInfo = new ConnectionInfo();
        static bool HasError = false;

        #region config
        
        static string OUTPUT_VERBOSITY
        {
            get
            {
                string thisValue = System.Configuration.ConfigurationManager.AppSettings["OUTPUT_VERBOSITY"];
                if (thisValue == "")
                {
                    thisValue = "NONE";
                }
                // TODO validation & error checking
                return thisValue;
            }
        }

        static string MAILER_PROFILE
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["MAILER_PROFILE"];
            }
        }

        static string RECIPIENTS
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["RECIPIENTS"];
            }
        }

        static string REPORT_NAME
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["REPORT_NAME"];
            }
        }
        static string OUTPUT_PATH
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["OUTPUT_PATH"];
            }
        }
        
        static string TargetReportServer
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["REPORT_SERVER"];
            }
        }
        static string TargetReportDatabase
        {
            get {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["REPORT_DATABASE"];
            }
        }

        static string TargetMailerServer
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["MAILER_SERVER"];
            }
        }
        static string TargetMailerDatabase
        {
            get
            {
                // TODO validation & error checking
                return System.Configuration.ConfigurationManager.AppSettings["MAILER_DATABASE"];
            }
        }
        #endregion
        static void VerboseWriteline(string s)
        {
            if (OUTPUT_VERBOSITY != "NONE")
            {
                Console.WriteLine(s);
            }
        }
            static void crAssignConnectionInfo()
        {
            crConnectionInfo.UserID = "";
            crConnectionInfo.Password = "";
            crConnectionInfo.DatabaseName = TargetReportDatabase;
            crConnectionInfo.ServerName = TargetReportServer;
            crConnectionInfo.IntegratedSecurity = true; // in case the report was saved qith SQL authentication, switch to Integrated
        }

        static void SetSubreportLoginInfo(CrystalDecisions.CrystalReports.Engine.Sections objSections)
        {
            foreach (Section section in objSections)
            {
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    SubreportObject crSubreportObject;
                    switch (reportObject.Kind)
                    {
                        case ReportObjectKind.SubreportObject:
                            crSubreportObject = (SubreportObject)reportObject;
                            ReportDocument subRepDoc = crSubreportObject.OpenSubreport(crSubreportObject.SubreportName);
                            if (subRepDoc.ReportDefinition.Sections.Count > 0)
                            {
                                SetSubreportLoginInfo(subRepDoc.ReportDefinition.Sections);
                            }
                            Tables crTables = subRepDoc.Database.Tables;
                            foreach (Table table in crTables)
                            {
                                TableLogOnInfo tableLogOnInfo = new TableLogOnInfo();
                                tableLogOnInfo.ConnectionInfo.UserID = crConnectionInfo.UserID;
                                tableLogOnInfo.ConnectionInfo.Password = crConnectionInfo.Password;
                                tableLogOnInfo.ConnectionInfo.DatabaseName = crConnectionInfo.DatabaseName;
                                tableLogOnInfo.ConnectionInfo.ServerName = crConnectionInfo.ServerName;
                                tableLogOnInfo.ConnectionInfo.IntegratedSecurity = crConnectionInfo.IntegratedSecurity;

                                table.ApplyLogOnInfo(tableLogOnInfo);
                            }
                            break;
                        case ReportObjectKind.FieldObject:
                        case ReportObjectKind.TextObject:
                        case ReportObjectKind.LineObject:
                        case ReportObjectKind.BoxObject:
                        case ReportObjectKind.PictureObject:
                        case ReportObjectKind.ChartObject:
                        case ReportObjectKind.CrossTabObject:
                        case ReportObjectKind.BlobFieldObject:
                        case ReportObjectKind.MapObject:
                        case ReportObjectKind.OlapGridObject:
                        case ReportObjectKind.FieldHeadingObject:
                        case ReportObjectKind.FlashObject:
                        default:
                            // none of the other objects need to have login assigned
                            break;
                    }
                }
            }
        }

        static void SetCrystalDocumentLogon()
        {
            if (HasError)
            {
                VerboseWriteline("Skipping Set login after error was encountered.");
                return;
            }

            crAssignConnectionInfo();
            TableLogOnInfo crTableLogonInfo = new TableLogOnInfo();
            foreach (Table crTable in crReportDocument.Database.Tables)
            {
                try
                {
                    crConnectionInfo.Type = crTable.LogOnInfo.ConnectionInfo.Type;
                    crTableLogonInfo.ConnectionInfo = crConnectionInfo;
                    crTableLogonInfo.ReportName = crTable.LogOnInfo.ReportName;
                    crTableLogonInfo.TableName = crTable.LogOnInfo.TableName;

                    crTable.ApplyLogOnInfo(crTableLogonInfo);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error during SetCrystalDocumentLogon " + ex.Message);
                    throw;
                }
                SetSubreportLoginInfo(crReportDocument.ReportDefinition.Sections);
            }
        }

        static bool crParameterInUse(String strParameterName)
        {
            if (crReportDocument.ParameterFields[strParameterName] == null || crReportDocument.ParameterFields[strParameterName].ToString() == "")
            {
                return false;
            }
            else
            {
                return crReportDocument.ParameterFields[strParameterName].ParameterFieldUsage2 != ParameterFieldUsage2.NotInUse;
            }
        }
        static void crAssignDefaultParameters()
        {
            if (HasError)
            {
                VerboseWriteline("Skipping crAssignDefaultParameters after error was encountered.");
                return;
            }

            foreach (ParameterField item in crReportDocument.ParameterFields)
            {
                // Console.WriteLine("{0} In use? {1}", item.Name, crParameterInUse(item.Name));

                if (crParameterInUse(item.Name))
                {
                    if (item.DefaultValues.Count == 0)
                    {
                        VerboseWriteline("Warning: no default value for " + item.Name + " assigned a value of empty string.");

                        crReportDocument.SetParameterValue(item.Name, "");
                    }
                    else
                    {
                        VerboseWriteline("Assigned default value of " + item.DefaultValues[0] + " for " + item.Name + "." );
                        crReportDocument.SetParameterValue(item.Name, ((CrystalDecisions.Shared.ParameterDiscreteValue)item.DefaultValues[0]).Value);
                    }
                }
                else
                {
                    VerboseWriteline("Skipping " + item.Name + ", assuming this is a linked subreport parameter for " + item.ReportName + ".");

                }
            }
        }
        static void crOpenReportFile()
        {
            VerboseWriteline("Opening Crystal Report file " + crReportName + " ");
            crReportDocument.Load(crReportName);
            try
            {
                crReportDocument.Load(crReportName);
            }
            catch (Exception ex)
            {
                HasError = true;
                Console.WriteLine("Check [List Folder / Read Data] permissions on C:\\Windows\\temp\\");
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        //***********************************************************************************************************************************
        public static string TrustedConnectionString(string server, string database)
        //***********************************************************************************************************************************
        {
            // see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemDataSqlClientSqlConnectionClassConnectionStringTopic.asp
            //
            // there is some debate as to whether the Oledb provider is indeed faster than the native client!
            //  
            string appname = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            string computername = Environment.MachineName.ToString();
            return "Workstation ID=" + computername + "_" + appname + ";" +
                   "packet size=8192;" +
                   "Persist Security Info=false;" +
                   "Server=" + server + ";" +
                   "Database=" + database + ";" +
                   "Trusted_Connection=true; " +
                   // "Network Library=dbmssocn;" +
                   "Pooling=True; " +
                   "Enlist=True; " +
                   "Connection Lifetime=14400; " +
                   "Max Pool Size=20; Min Pool Size=0";
        }

        static bool TestConnectionString(string connectionString)
        {
            string strSQL = "select suser_sname()";
            string thisUserSQL = "";
            try
            {
                thisUserSQL = (string)SqlHelper.ExecuteScalar(connectionString, CommandType.Text, strSQL);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error when attempting to determine SQL user context for TargetServer = {0}, Target Database = {1}; {2}", TargetReportServer, TargetReportDatabase, ex.Message);
                HasError = true;
                return false;
            }
            return true;
        }
        /// <summary>
        /// Send report via SQL sp_send_dbmail (SQL service account needs read permissions to reports directory)
        /// </summary>
        static void SendReport()
        {
            if (HasError)
            {
                VerboseWriteline("Skipping SendReport after error was encountered.");
                return;
            }

            string strSQL = "msdb.dbo.sp_send_dbmail";
            SqlParameter[] sqlParams = new SqlParameter[]
            {
                 new SqlParameter("@profile_name",MAILER_PROFILE),
                 new SqlParameter("@recipients",RECIPIENTS),
                 new SqlParameter("@subject","COVID19 Situation Status Report for " + DateTime.Now.ToLongDateString() + " at " + DateTime.Now.ToLongTimeString()  ),
                 new SqlParameter("@body","Please see attached PDF file."),
                 new SqlParameter("@file_attachments", OutputFileWithPath),
            };
            try
            {
                // we'll be running this app on a different server that the data warehouse, assumed to have SQL and Crystal Reports locally
                string thisConnectionString = TrustedConnectionString(TargetMailerServer, TargetMailerDatabase);
                TestConnectionString(thisConnectionString);
                SqlHelper.ExecuteNonQuery(thisConnectionString, CommandType.StoredProcedure, strSQL, sqlParams);
                Console.WriteLine("Messge sent.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error when attempting to call sp_send_dbmail: " + ex.Message);
                HasError = true;
                return;
            }
        }

        static void SaveReport()
        {
            if (HasError)
            {
                VerboseWriteline("Skipping SaveReport after error was encountered.");
                return;
            }

            try
            {
                crReportDocument.ExportToDisk(ExportFormatType.PortableDocFormat, OutputFileWithPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error when saving PDF file: " + ex.Message);
                HasError = true;
            }
        }

        static bool init(string[] args)
        {
            string crReportNameWithoutExtension;
            string thisReportName;
            crReportDocument = new ReportDocument();

            thisReportName = REPORT_NAME;

            if (Path.GetExtension(thisReportName).ToLower() != ".rpt")
            {
                thisReportName += ".rpt";
            }
            crReportName = Path.GetFullPath(thisReportName);
            crReportNameWithoutExtension = Path.GetFileNameWithoutExtension(thisReportName);
            String FileDateStampSuffix = DateTime.Now.Year.ToString().PadLeft(4, '0') + "_" +
                                         DateTime.Now.Month.ToString().PadLeft(2, '0') + "_" +
                                         DateTime.Now.Day.ToString().PadLeft(2, '0');
            VerboseWriteline("Looking for " + OUTPUT_PATH);
            if (!Directory.Exists(OUTPUT_PATH))
            {
                Directory.CreateDirectory(OUTPUT_PATH);
            }
            OutputFileWithPath = OUTPUT_PATH + crReportNameWithoutExtension + "_" + FileDateStampSuffix + ".pdf";

            return true;
        }

        /// <summary>
        /// main app
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            if (init(args))
            {
                crOpenReportFile();
                crAssignDefaultParameters();
                SetCrystalDocumentLogon();
                SaveReport();
                SendReport();
            }
            else
            {
                HasError = true;
                Console.WriteLine("Error during initialization; abort");
            }
            if (HasError)
            {
                Environment.ExitCode = -1;
            }
            System.Threading.Thread.Sleep(5000);
        }
    }
}

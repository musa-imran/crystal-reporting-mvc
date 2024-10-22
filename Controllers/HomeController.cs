using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.ReportAppServer.DataDefModel;
using CrystalDecisions.Shared;
using MvcAppTest.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Drawing.Printing;
using ConnectionInfo = CrystalDecisions.Shared.ConnectionInfo;
using Tables = CrystalDecisions.CrystalReports.Engine.Tables;

namespace MvcAppTest.Controllers
{
    public class HomeController : Controller
    {
        private AzureDBEntities db = new AzureDBEntities();
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }










        //======================================================================= Controller Method ===========================================================

        public ActionResult GenerateReport(string viewReport, string downloadReport, string printReport, string telerikViewer)
        {
            ReportDocument reportDocument = new ReportDocument();

            try
            {
                string reportPath = Server.MapPath("~/Reports/productReport.rpt");
                FileInfo fileInfo = new FileInfo(reportPath);

                if (!fileInfo.Exists)
                {
                    return Content("Error: Report file does not exist at path: " + reportPath);
                }

                reportDocument.Load(reportPath);

                //DateTime fdate = DateTime.ParseExact("2024-10-01", "yyyy-MM-dd", null);
                //DateTime tdate = DateTime.ParseExact("2024-10-05", "yyyy-MM-dd", null);
                
                //reportDocument.SetParameterValue("fdt", fdate);
                //reportDocument.SetParameterValue("tdt", tdate);
                //reportDocument.RecordSelectionFormula = "{Product.dt}>=#" + fdate.ToString("yyyy-MM-dd") + "# and {Product.dt}<=#" + tdate.ToString("yyyy-MM-dd") + "#";

                SetReportSource(ref reportDocument);

                if (!string.IsNullOrEmpty(viewReport))
                {
                    Stream stream = reportDocument.ExportToStream(ExportFormatType.PortableDocFormat);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream, "application/pdf");
                }

                else if (!string.IsNullOrEmpty(telerikViewer))
                {
                    Stream stream = reportDocument.ExportToStream(ExportFormatType.PortableDocFormat);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream, "application/pdf");
                }

                else if (!string.IsNullOrEmpty(downloadReport))
                {
                    Stream stream = reportDocument.ExportToStream(ExportFormatType.PortableDocFormat);
                    stream.Seek(0, SeekOrigin.Begin);
                    return File(stream, "application/pdf", "YourReport.pdf");
                }

                else if (!string.IsNullOrEmpty(printReport))
                {
                    string printerName = "";
                    reportDocument.PrintOptions.PrinterName = printerName;
                    reportDocument.PrintToPrinter(1, false, 0, 0);
                    return Content("Print Completed!");

                    //return Content("Printer Not Found");
                }
                else
                {
                    return HttpNotFound();
                }
            }
            catch (Exception ex)
            {
                return Content("Error generating report: " + ex.Message);
            }
            finally
            {
                reportDocument.Close();
                reportDocument.Dispose();
            }
        }



        //======================================================================= Crystal Report Source Reference ===========================================================
        private void SetReportSource(ref ReportDocument reportDocument)
        {

            ConnectionInfo connectionInfo = new ConnectionInfo
            {
                ServerName = "SDPC9",
                DatabaseName = "BlazorDB",
                UserID = "sa",
                Password = "123456"
            };

            Tables tables = reportDocument.Database.Tables;
            foreach (CrystalDecisions.CrystalReports.Engine.Table table in tables)
            {
                TableLogOnInfo tableLogOnInfo = table.LogOnInfo;
                tableLogOnInfo.ConnectionInfo = connectionInfo;
                table.ApplyLogOnInfo(tableLogOnInfo);
            }

            foreach (Section section in reportDocument.ReportDefinition.Sections)
            {
                foreach (ReportObject reportObject in section.ReportObjects)
                {
                    if (reportObject.Kind == ReportObjectKind.SubreportObject)
                    {
                        SubreportObject subreport = (SubreportObject)reportObject;
                        ReportDocument subReportDocument = subreport.OpenSubreport(subreport.SubreportName);
                        foreach (CrystalDecisions.CrystalReports.Engine.Table subTable in subReportDocument.Database.Tables)
                        {
                            TableLogOnInfo subTableLogOnInfo = subTable.LogOnInfo;
                            subTableLogOnInfo.ConnectionInfo = connectionInfo;
                            subTable.ApplyLogOnInfo(subTableLogOnInfo);
                        }
                    }
                }
            }
        }





    }

}
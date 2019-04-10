using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.Text;
using System.Net.Mail;
using System.IO;
using CrystalDecisions.Shared;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Web;
using Automate.Utilities;



public partial class Report : System.Web.UI.Page
{
    ReportDocument cryRpt = new ReportDocument();
	string test = string.empty;
    protected void Page_Init(object sender, EventArgs e)
    {
        showreport();

        //mail();
        //pdf1();
        
    }

    private void showreport()
    {
        cryRpt.Load(Server.MapPath("~/Report/Report1.rpt"));
        cryRpt.SetDatabaseLogon("btpconline", "admin@local", "TSWCLB-DB-SRV", "TSWDATA_ClientCustom");
        CrystalReportViewer1.ReportSource = cryRpt;


        print();
    }

    private void mail()
    {
       



        MailMessage Msg = new MailMessage();
        Msg.Subject = "testemail";
        Msg.BodyEncoding = Encoding.GetEncoding("UTF-8");
        Msg.SubjectEncoding = Encoding.GetEncoding("UTF-8");
        Msg.From = new MailAddress("noreply@angsanavacationclub.com", "Angsana Vacation Club");
        Msg.To.Add(new MailAddress("eakkaphols@lagunaphuket.com", "eakkaphol"));
      

        Msg.Body = "Quickly please!!!!";
        Msg.IsBodyHtml = true;
        ReportDocument rpt = new ReportDocument();
        rpt.Load(Server.MapPath("~/Report/Report2.rpt"));
        Msg.Attachments.Add(new Attachment(rpt.ExportToStream(ExportFormatType.PortableDocFormat),  "CFMLetter.pdf"));
        //System.IO.Stream oStream = null;
        //byte[] byteArray = null;
        //oStream = rpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
        //byteArray = new byte[oStream.Length];
        //oStream.Read(byteArray, 0, Convert.ToInt32(oStream.Length - 1));
        //Response.ClearContent();
       // Response.ClearHeaders();
       // Response.ContentType = "application/pdf";
       // Response.BinaryWrite(byteArray);
        //Response.Flush();
        //Response.Close();

        //oStream = (MemoryStream)rpt.ExportToStream(ExportFormatType.PortableDocFormat);

       // Response.ContentType = "application/pdf";
        //Msg.Attachments.Add(new Attachment(oStream, "-CFMLetter.pdf"));

       
        string factsheet = Server.MapPath("~/pdf/test.pdf");
        Attachment attachment = new Attachment(factsheet);
        attachment.Name = "Factsheet.pdf";
        Msg.Attachments.Add(attachment);
        Helpers.SendHTMLMail(Msg);
    }

    private void pdf1()
    {
        ReportDocument rpt = new ReportDocument();
        rpt.Load(Server.MapPath("~/Report/Report2.rpt"));
        System.IO.Stream oStream = null;
        byte[] byteArray = null;
        oStream = rpt.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
        byteArray = new byte[oStream.Length];
        oStream.Read(byteArray, 0, Convert.ToInt32(oStream.Length - 1));
        Response.ClearContent();
        Response.ClearHeaders();
        Response.ContentType = "application/pdf";
        Response.BinaryWrite(byteArray);
        Response.Flush();
        Response.Close();
        rpt.Close();
        rpt.Dispose();
    }

    private void print()
    {
        //CrystalReportViewer1.ReportSource = Session["ReportSource"];
        int exportFormatFlags = (int)(ViewerExportFormats.PdfFormat |
                                ViewerExportFormats.ExcelFormat |
                                ViewerExportFormats.ExcelRecordFormat |
                                ViewerExportFormats.XLSXFormat);
        CrystalReportViewer1.AllowedExportFormats = exportFormatFlags;
        CrystalReportViewer1.Visible = true;

    }
    protected void Page_UnLoad(object sender, EventArgs e)
    {
        CrystalReportViewer1.Dispose();
        CrystalReportViewer1 = null;
        cryRpt.Close(); // I can't remember if this is part of the reportDucment class
        cryRpt.Dispose();
        cryRpt = null;
        GC.Collect(); // I forgot this line
    }


}
using Consignment.Model;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.qrcode;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing.Imaging;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ZXing;


namespace Consignment.ConsignmentPanel
{
    public partial class ReportViewer : System.Web.UI.Page
    {
        #region [properties]
        ArrayList ParameterName, ParameterValue;
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
        connect con = new connect();
        common cm = new common();
        #endregion

        #region[PageLoad]
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString["BookId"] != null)
            {

            }
            else
            {
                Response.Redirect("AddBookLr.aspx");
            }
        }
        #endregion

        #region[Function]
        private void GenerateLRPdf(string BookId, string FileName, string Reporttxt)
        {
            string ReportNo = string.Empty;
            int LoopCount = 0;
            DataTable dt = con.GetDatatable("Select * from Vw_RptLRBook where Bookid='" + BookId + "'");
            if (dt.Rows[0]["MType"].ToString() == "B")
            {
                int GroupId = Convert.ToInt32((con.GetDatatable("Select * from Vw_RptLRBook where Bookid='" + BookId + "'")).Rows[0]["GroupId"]);
                dt = con.GetDatatable("Select * from Vw_RptLRBook where groupid='" + GroupId + "'");
            }
            string spName = "";
            string StrBase64 = "";
            string CompanyImageStr = string.Empty;
            string CompanyName = string.Empty;
            string CompanyGSTIN = string.Empty;
            string CompanyAddress1 = string.Empty;
            string CompanyAddress2 = string.Empty;
            string CompanyPhone1 = string.Empty;
            string CompanyPhone2 = string.Empty;
            string CompanyFax = string.Empty;
            string CompanyEmail = string.Empty;
            string CompanyWebsite = string.Empty;
            string DocumentNo = string.Empty;
            string EwayDetails = string.Empty;
            string InvoiceDetails = string.Empty;

            DataTable dtCompany = con.GetDatatable("select * from CompanyMaster");
            CompanyAddress1 = dtCompany.Rows[0]["Address1"].ToString();
            CompanyAddress2 = dtCompany.Rows[0]["Address2"].ToString();
            CompanyPhone1 = dtCompany.Rows[0]["Phone1"].ToString();
            CompanyPhone2 = dtCompany.Rows[0]["Phone2"].ToString();
           
            if (dtCompany.Rows[0]["Fax"].ToString() != "")
            {
                CompanyFax = ", Fax : " + dtCompany.Rows[0]["Fax"].ToString();
            }
            if (dtCompany.Rows[0]["Website"].ToString() != "")
            {
                CompanyWebsite = ", Web : " + dtCompany.Rows[0]["Website"].ToString();
            }
            CompanyEmail = dtCompany.Rows[0]["Email1"].ToString();
            
            CompanyImageStr = dtCompany.Rows[0]["ImageStr"].ToString();
            CompanyName = dtCompany.Rows[0]["CompanyName"].ToString();
            CompanyGSTIN = "GSTIN : " + dtCompany.Rows[0]["GSTIN"].ToString() + "                  CIN:-" + dtCompany.Rows[0]["CIN"].ToString() + "                        PAN NO:" + dtCompany.Rows[0]["PAN"].ToString() + "                         FSSAI L No : " + dtCompany.Rows[0]["FSSAI"].ToString();
            string CompanyGSTIN1 = "GSTIN : " + dtCompany.Rows[0]["GSTIN"].ToString() + ",  PAN NO:" + dtCompany.Rows[0]["PAN"].ToString();

            string strPdfPath = FolderCheck("App_Content/Report");
            FileStream fs = new FileStream(strPdfPath + "\\" + FileName, FileMode.Create, FileAccess.Write, FileShare.None);
            Document document = new Document((iTextSharp.text.PageSize.A4));
            document.SetMargins(30.0f, 30.0f, 30.0f, 30.0f);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();
            BaseFont bf = BaseFont.CreateFont(@"C:\Windows\Fonts\seguisym.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            float[] columnWidths = { 190, 190, 190, 190, 190, 190, 190, 190 };
            if (Reporttxt == "") { LoopCount = 4; } else { LoopCount = 1; }
            for (int i = 1; i <= LoopCount; i++)
            {
                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    DocumentNo = "";
                    spName = "Proc_DocumentDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAllForPrint");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable DocumentDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (DocumentDt.Rows.Count > 0)
                    {
                        for (int d = 0; d < DocumentDt.Rows.Count; d++)
                        {
                            DocumentNo += " [" + DocumentDt.Rows[d]["DocumentNo"].ToString() + "] ";
                        }
                    }

                    EwayDetails = "EwayBill ";
                    spName = "Proc_EwayBillDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAllforPrint");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable EwayDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (EwayDt.Rows.Count > 0)
                    {
                        for (int e = 0; e < EwayDt.Rows.Count; e++)
                        {
                            EwayDetails += " [No.:" + EwayDt.Rows[e]["EwayBillNo"].ToString() + ",Date:" + EwayDt.Rows[e]["BillDate"].ToString() + ",Upto:" + EwayDt.Rows[e]["BillUptoDate"].ToString() + "] ";
                        }
                    }

                    InvoiceDetails = "Invoice ";
                    spName = "Proc_InvoiceDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAllforPrint");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable InvoiceDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (InvoiceDt.Rows.Count > 0)
                    {
                        for (int inv = 0; inv < InvoiceDt.Rows.Count; inv++)
                        {
                            InvoiceDetails += " [No.:" + InvoiceDt.Rows[inv]["InvoiceNo"].ToString() + ",Date:" + InvoiceDt.Rows[inv]["InvoiceDate"].ToString() + ",Value:" + InvoiceDt.Rows[inv]["InvoiceValue"].ToString() + "] ";
                        }
                    }

                    StrBase64 = GenerateBarCode(dt.Rows[z]["BookLrNo"].ToString());
                    if (Reporttxt == "")
                    {
                        if (i == 1) { ReportNo = "CONSIGNOR COPY"; }
                        else if (i == 2) { ReportNo = "CONSIGNEE COPY"; document.NewPage(); }
                        else if (i == 3) { ReportNo = "DRIVER COPY"; document.NewPage(); }
                        else if (i == 4) { ReportNo = "ACCOUNT COPY"; }
                    }
                    else
                    {
                        ReportNo = Reporttxt;
                    }
                    PdfPTable table = null;
                    table = new PdfPTable(columnWidths);
                    table.WidthPercentage = 100;
                    iTextSharp.text.Font fontB = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font font = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font ELR = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font Vayudoot = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font Copy = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font VayudootAddress = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font AtOwnerRisk = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font notice = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.RED));
                    iTextSharp.text.Font INSURANCE = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font caution = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font TheConsignmentCovered = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.RED));
                    iTextSharp.text.Font SubjectToPune = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font SubjectToPune1 = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font AddressOfDelivery = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font DoorDelivery = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font TheConsignerHasStated = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font PolicyNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font VRCGSTIN = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtConName = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConName = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConAddress = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConGSTIN = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtConNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font ConNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtVehicleNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 12, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font VechicleNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font FromTo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font VechicleSize = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font Sign = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font LabelBold = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font Label = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font LabelT = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 5, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font Date = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font CheckboxText = new iTextSharp.text.Font(bf, 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
                    string TermsAndConditions =
 "ALL CONSIGNMENTS ARE ACCEPTED FOR TRANSPORTATION SUBJECT TO OUR TERMS AND CONDITIONS EXHIBITED HERE AS UNDER.\n\n" +

 "1. All Consignment are Subject to Mumbai Jurisdiction.\n" + 
 
 "2. Packing nature of contents and value of the consignment are not known to us." +
 "We carry the goods as declared by the Consignor.\n" +

 "3. Consignor will be liable for the damage caused by sub-standard packing." +
 "Consignor will furnish necessary documents viz. Challan, Invoice, GST Way Bill along with Part A & B, Insurance Copy & Other Documents of Entry to the Driver.\n" +

 "4. Carrier will not be responsible if any goods declared contraband and confiscated by the Government.\n" +

 "5. Goods to be unloaded according to the \"Maric\". If wrong goods are unloaded, Carrier will not be responsible." +
 "Loss made by wrong goods to be borne by the either consignor or consignee.\n" +

 "6. All consignments will be accepted at \"Owner’s Risk\" for booking, unless otherwise expressly agreed in writing at Carrier’s Risk.\n" +

 "7. Goods will be directly delivered to the person/firm/company/institution whose name is written in the consignment note, without asking original consignee copy.\n" +

 "8. Insurance for transit risk of consignments will have to be done by the consignor/consignee.\n" +

 "9. Carrier will not be liable to dispose the goods lying in the Godown undelivered within 30 days of arrival after Notice in writing to the consignor / consignee." +
 "Demurrage will be charged.\n" +

 "10. Carrier do not guarantee delivery of the consignment and shall not be liable for delay due to breakdown, accident of the vehicle which beyond the control of the driver.\n" +

 "11. In case of loss of consignment copy delivery of the consignment will be made against indemnity bond in all respects.\n" +

 "12. Consignments of Liquid, viz. Ghee, Oil, Paints, Fragile Goods, Edibles and Fruits, vegetables, fragile goods shall be booked at \"Owner’s Risk\"." +
 "Carrier will not be liable for any type of loss.\n" +

 "13. This consignment is approved only after filling the goods from consignor place against our receipt.\n" +

 "14. ALL GOODS NEED TO BE INSURED BY CONSIGNOR OR CONSIGNEE, TILL THE GOODS ARE DELIVERED TO THE PARTY/CONSIGNEE.\n" +

 "15. Detention charge, Loading/Unloading and other incidental charges will be Responsibility of Consignor/Consignee.\n" +

 "16. Once Vehicle Reported at Loading point extension of E-Way Bill Required.";



                    PdfPCell cell = new PdfPCell();

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 8;
                    cell.Padding = 0;
                    cell.FixedHeight = 6f;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    byte[] imageBytes = Convert.FromBase64String(CompanyImageStr);
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageBytes);
                    float fixedHeight = 60f;
                    float ratio = jpg.Width / jpg.Height;
                    jpg.ScaleAbsolute(fixedHeight * ratio, fixedHeight);
                    cell = new PdfPCell(jpg);
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 4;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(CompanyName, Vayudoot));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 0;
                    cell.Border = 0;
                    table.AddCell(cell);

                    imageBytes = Convert.FromBase64String(StrBase64);
                    jpg = iTextSharp.text.Image.GetInstance(imageBytes);
                    jpg.ScaleAbsolute(130f, 29.0f);
                    cell = new PdfPCell(jpg);
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 2;
                    cell.FixedHeight = 40f;   // <-- fixed box height

                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);




                    cell = new PdfPCell(new Phrase("(CARGO MOVERS & TRANSPORT CONTRACTORS) \n Corporate Office : " + CompanyAddress1 + " \n " + CompanyAddress2 + " \n " + CompanyPhone1 + ", " + CompanyPhone2 + CompanyFax + " \n Email : " + CompanyEmail +  CompanyWebsite + " \n " + CompanyGSTIN1, VayudootAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Rowspan = 2;
                    cell.PaddingBottom = 4;
                    cell.Border = 0;
                    table.AddCell(cell);


                    if (dt.Rows[0]["MType"].ToString() == "B")
                    {
                        cell = new PdfPCell(new Phrase("Consignment No (" + (z + 1) + "/" + dt.Rows.Count.ToString() + ") :", txtConNo));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Colspan = 2;
                        cell.Rowspan = 1;
                        cell.PaddingBottom = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 1;
                        table.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase("Consignment No :", txtConNo));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Colspan = 2;
                        cell.Rowspan = 1;
                        cell.PaddingBottom = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 1;
                        table.AddCell(cell);
                    }




                    cell = new PdfPCell(new Phrase(ReportNo, txtVehicleNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Rowspan = 1;
                    cell.PaddingBottom = 4;
                    cell.Border = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookLrNo"].ToString(), ConNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.PaddingBottom = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);


                    cell = new PdfPCell(new Phrase("Consignor  : " + dt.Rows[z]["ConsignorName"].ToString(), LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Date : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Booking Office : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 4;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsignorAdd"].ToString() + "," + dt.Rows[z]["ConsignorState"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookDate1"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dtCompany.Rows[0]["BookingOfficeAdd"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 4;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase("Vehicle Size :", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("From : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["SourceFrom"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Address Of The Delivery: ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);




                    cell = new PdfPCell(new Phrase("GSTIN : " + dt.Rows[z]["ConsignorGST"].ToString(), LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookCustSize"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("To : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ToDestination"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookType"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);




                    cell = new PdfPCell(new Phrase("Consignee  : " + dt.Rows[z]["ConsigneeName"].ToString(), LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Vehicle No. : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Billed At : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Rate", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Rs.", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsigneeAdd"].ToString() + "," + dt.Rows[z]["ConsigneeState"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["VchNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                   

                    cell = new PdfPCell(new Phrase("TBB", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Material Type : ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ProdName"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("To be Billed To :", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);


                    if (dt.Rows[z]["BookBillTo"].ToString() == "TOPAY")
                    {
                        cell = new PdfPCell(new Phrase( dt.Rows[z]["Consigneename"].ToString(), font));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        cell.Colspan = 1;
                        cell.Rowspan = 2;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase( dt.Rows[z]["BookBillTo"].ToString(), font));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                        cell.Colspan = 1;
                        cell.Rowspan = 2;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 0;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }




                    cell = new PdfPCell(new Phrase("Freight Basic", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("GSTIN : " + dt.Rows[z]["ConsigneeGST"].ToString(), LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Loading", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Invoice No ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase("Date ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Qty ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("E-way Bill No", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Valid Upto", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("P.O./Doc No.", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Unloading", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("  ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    int LoopDigit = 0;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);









                    cell = new PdfPCell(new Phrase("FOV Charges ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("  ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("ODA Charge ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("To", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);


                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Local Charges", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Be", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);


                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("D.D. Charges", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Billed", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("SR Charges", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("At", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Doc. Charges", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Bombay", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Other Charges", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Total", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("TDS (-)", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);


                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("SGST 0%", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("CGST 0%", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);




                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("IGST 0%", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Grand Total", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 1;
                    cell.Rowspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);


                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    LoopDigit += 1;
                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(InvoiceDt.Rows[LoopDigit]["InvoiceUnit"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["EwayBillNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDt.Rows[LoopDigit]["BillUptoDate"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(DocumentDt.Rows[LoopDigit]["DocumentNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("CAUTION", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("No. of Packages", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookQty"].ToString() + " " + dt.Rows[z]["UnitName"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Seal No.", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["SealNo"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Other (Remarks)", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                   

                    cell = new PdfPCell(new Phrase("This Consignment Will Not Be Detained,\r\nDiverted, Re-routed Or Re-booked Without\r\nConsignee’s Written Permission.\r\nWill Be Delivered At The Destination.", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("GST Payable By:", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 4;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookRemark1"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    string gstPay = dt.Rows[z]["GSTPay"].ToString().Trim();

                    string consignor = "☐";
                    string consignee = "☐";
                    string transporter = "☐";
                    string na = "☐";

                    if (gstPay.Equals("Consignor", StringComparison.OrdinalIgnoreCase))
                    {
                        consignor = "☑";
                    }
                    else if (gstPay.Equals("Consignee", StringComparison.OrdinalIgnoreCase))
                    {
                        consignee = "☑";
                    }
                    else if (gstPay.Equals("Transporter", StringComparison.OrdinalIgnoreCase))
                    {
                        transporter = "☑";
                    }
                    else if (gstPay.Equals("NA", StringComparison.OrdinalIgnoreCase)
                          || gstPay.Equals("Not Applicable", StringComparison.OrdinalIgnoreCase))
                    {
                        na = "☑";
                    }

                    string gstText = consignor + " Consignor   "
                                   + consignee + " Consignee   "
                                   + transporter + " Transporter   "
                                   + na + " Not Applicable  (As Per Rules)";

                    cell = new PdfPCell(new Phrase(gstText, CheckboxText));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 4;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;

                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase("Total Invoice Value:", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("As Per Attached Invoices", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Risk Type:", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Owner's Risk", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                   

                    cell = new PdfPCell(new Phrase("Remarks:", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookRemarks"].ToString(), Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 3;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Generated By & Sign:", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookLoading"].ToString(), LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 1;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);


                   

                    cell = new PdfPCell(new Phrase("Terms & Conditions ", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 8;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(TermsAndConditions, LabelT));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 8;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Acknowledgement:", LabelBold));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_MIDDLE;
                    cell.Colspan = 8;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Goods Received In Good Condition With Assign Quantity", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Colspan = 8;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("[Consignee / Receiver Rubber Stamp & Authorised Signature]", Label));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Colspan = 8;
                    cell.Rowspan = 1;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    document.Add(table);

                    document.NewPage();
                }


            }

            document.Close();
            document.Dispose();
            writer.Close();
            writer.Dispose();
            fs.Close();
            fs.Dispose();
        }



        private void PrintAllReport_iText(string BookId, string FileName, string Reporttxt)
        {
            string ReportNo = string.Empty;
            int LoopCount = 0;
            DataTable dt = con.GetDatatable("Select * from Vw_RptLRBook where Bookid='" + BookId + "'");
            if (dt.Rows[0]["MType"].ToString() == "B")
            {
                int GroupId = Convert.ToInt32((con.GetDatatable("Select * from Vw_RptLRBook where Bookid='" + BookId + "'")).Rows[0]["GroupId"]);
                dt = con.GetDatatable("Select * from Vw_RptLRBook where groupid='" + GroupId + "'");
            }
            string spName = "";
            string StrBase64 = "";
            string CompanyImageStr = string.Empty;
            string CompanyName = string.Empty;
            string CompanyGSTIN = string.Empty;
            string CompanyAddress1 = string.Empty;
            string CompanyAddress2 = string.Empty;
            string CompanyPhone1 = string.Empty;
            string CompanyPhone2 = string.Empty;
            string CompanyFax = string.Empty;
            string CompanyEmail = string.Empty;
            string CompanyWebsite = string.Empty;
            string DocumentNo = string.Empty;
            string EwayDetails = string.Empty;
            string InvoiceDetails = string.Empty;

            DataTable dtCompany = con.GetDatatable("select * from CompanyMaster");
            CompanyAddress1 = dtCompany.Rows[0]["Address1"].ToString();
            CompanyAddress2 = dtCompany.Rows[0]["Address2"].ToString();
            CompanyPhone1 = dtCompany.Rows[0]["Phone1"].ToString();
            CompanyPhone2 = dtCompany.Rows[0]["Phone2"].ToString();
            CompanyFax = dtCompany.Rows[0]["Fax"].ToString();
            CompanyEmail = dtCompany.Rows[0]["Email1"].ToString();
            CompanyWebsite = dtCompany.Rows[0]["Website"].ToString();
            CompanyImageStr = dtCompany.Rows[0]["ImageStr"].ToString();
            CompanyName = dtCompany.Rows[0]["CompanyName"].ToString();
            CompanyGSTIN = "GSTIN : " + dtCompany.Rows[0]["GSTIN"].ToString() + "                  CIN:-" + dtCompany.Rows[0]["CIN"].ToString() + "                        PAN NO:" + dtCompany.Rows[0]["PAN"].ToString() + "                         FSSAI L No : " + dtCompany.Rows[0]["FSSAI"].ToString();

            string strPdfPath = FolderCheck("App_Content/Report");
            FileStream fs = new FileStream(strPdfPath + "\\" + FileName, FileMode.Create, FileAccess.Write, FileShare.None);
            Document document = new Document((iTextSharp.text.PageSize.A4.Rotate()));
            document.SetMargins(30.0f, 30.0f, 30.0f, 30.0f);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();
            float[] columnWidths = { 98, 117, 123, 230, 273, 165, 101, 409 };
            if (Reporttxt == "") { LoopCount = 4; } else { LoopCount = 1; }
            for (int i = 1; i <= LoopCount; i++)
            {
                for (int z = 0; z < dt.Rows.Count; z++)
                {
                    DocumentNo = "";
                    spName = "Proc_DocumentDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAll");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable DocumentDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (DocumentDt.Rows.Count > 0)
                    {
                        for (int d = 0; d < DocumentDt.Rows.Count; d++)
                        {
                            DocumentNo += " [" + DocumentDt.Rows[d]["DocumentNo"].ToString() + "] ";
                        }
                    }

                    EwayDetails = "EwayBill ";
                    spName = "Proc_EwayBillDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAllforPrint");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable EwayDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (EwayDt.Rows.Count > 0)
                    {
                        for (int e = 0; e < EwayDt.Rows.Count; e++)
                        {
                            EwayDetails += " [No.:" + EwayDt.Rows[e]["EwayBillNo"].ToString() + ",Date:" + EwayDt.Rows[e]["BillDate"].ToString() + ",Upto:" + EwayDt.Rows[e]["BillUptoDate"].ToString() + "] ";
                        }
                    }

                    InvoiceDetails = "Invoice ";
                    spName = "Proc_InvoiceDetails";
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    ParameterName.Add("@Action"); ParameterValue.Add("GetAllforPrint");
                    ParameterName.Add("@LRNo"); ParameterValue.Add(dt.Rows[z]["BookLrNo"].ToString());
                    DataTable InvoiceDt = con.GetDataTableFromSP(spName, ParameterName, ParameterValue);
                    if (InvoiceDt.Rows.Count > 0)
                    {
                        for (int inv = 0; inv < InvoiceDt.Rows.Count; inv++)
                        {
                            InvoiceDetails += " [No.:" + InvoiceDt.Rows[inv]["InvoiceNo"].ToString() + ",Date:" + InvoiceDt.Rows[inv]["InvoiceDate"].ToString() + ",Value:" + InvoiceDt.Rows[inv]["InvoiceValue"].ToString() + "] ";
                        }
                    }

                    StrBase64 = GenerateBarCode(dt.Rows[z]["BookLrNo"].ToString());
                    if (Reporttxt == "")
                    {
                        if (i == 1) { ReportNo = "CONSIGNOR COPY"; }
                        else if (i == 2) { ReportNo = "CONSIGNEE COPY"; document.NewPage(); }
                        else if (i == 3) { ReportNo = "DRIVER COPY"; document.NewPage(); }
                        else if (i == 4) { ReportNo = "ACCOUNT COPY"; }
                    }
                    else
                    {
                        ReportNo = Reporttxt;
                    }
                    PdfPTable table = null;
                    table = new PdfPTable(columnWidths);
                    table.WidthPercentage = 100;
                    iTextSharp.text.Font fontB = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font font = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font ELR = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font Vayudoot = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font Copy = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font VayudootAddress = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font AtOwnerRisk = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 16, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font notice = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.RED));
                    iTextSharp.text.Font INSURANCE = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font caution = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font TheConsignmentCovered = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.RED));
                    iTextSharp.text.Font SubjectToPune = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font SubjectToPune1 = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font AddressOfDelivery = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font DoorDelivery = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font TheConsignerHasStated = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font PolicyNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font VRCGSTIN = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtConName = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConName = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConAddress = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font ConGSTIN = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtConNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font ConNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font txtVehicleNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 12, iTextSharp.text.Font.BOLD, BaseColor.RED));
                    iTextSharp.text.Font VechicleNo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font FromTo = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 13, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font VechicleSize = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK));
                    iTextSharp.text.Font Sign = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));
                    iTextSharp.text.Font Date = new iTextSharp.text.Font(FontFactory.GetFont("Calibri", 11, iTextSharp.text.Font.NORMAL, BaseColor.BLACK));

                    PdfPCell cell = new PdfPCell();
                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 7;
                    cell.Padding = 0;
                    cell.FixedHeight = 6f;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("E-LR", ELR));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    table.AddCell(cell);

                    byte[] imageBytes = Convert.FromBase64String(CompanyImageStr);
                    iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(imageBytes);
                    jpg.ScaleAbsolute(300f, 80f);
                    cell = new PdfPCell(jpg);
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 4;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(CompanyName, Vayudoot));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 0;
                    cell.Border = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("(CARGO MOVERS & TRANSPORT CONTRACTORS) \n Corporate Office : " + CompanyAddress1 + " \n " + CompanyAddress2 + " \n " + CompanyPhone1 + ", " + CompanyPhone2 + "  Fax : " + CompanyFax + " \n Email : " + CompanyEmail + ", Web : " + CompanyWebsite, VayudootAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Rowspan = 3;
                    cell.PaddingBottom = 4;
                    cell.Border = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("CAUTION", caution));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 10;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell();
                    cell.AddElement(new Phrase("Subject to pune Jurisdiction. ", SubjectToPune));
                    cell.AddElement(new Phrase("This consignment  will not be detained diverted,re-routed or re-booked without consignee Bank's  written permission.will be delivered at the destination.", SubjectToPune1));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(ReportNo, Copy));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.Border = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("AT OWNERS RISK", AtOwnerRisk));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.Border = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("NOTICE", notice));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Border = 0;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("INSURANCE", INSURANCE));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Border = 0;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    cell.Padding = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("THE CONSIGNMENT COVERED BY THIS SET OF SPECIAL LORRY Form shall be stored at the destination under the control of the Transport Operator and shall be delivered to the order of the consignee Bank Whose name is mention in Lorry Reciept.It will Under no Circumstances be dilivered to anyone with out the written authority from the consignee bank of its order endorsed on the consignee copy or on a separate Letter of authority", TheConsignmentCovered));
                    cell.Colspan = 4;
                    cell.Rowspan = 4;
                    cell.Padding = 10;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    cell.HorizontalAlignment = PdfPCell.ALIGN_JUSTIFIED;
                    table.AddCell(cell);

                    PdfContentByte cb = writer.DirectContent;
                    cb.Rectangle(330f, 432f, 6f, 6f);
                    cb.Rectangle(330f, 423f, 6f, 6f);
                    //if (i != 4) { cb = null; }
                    cell = new PdfPCell(new Phrase("This Consignor has stated that: \n   he has not insured the consignment or \n   he has insured the consignment company", TheConsignerHasStated));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 3;
                    cell.PaddingLeft = 10;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Address of the delivery", AddressOfDelivery));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 5;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Door Delivery", DoorDelivery));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Policy No :                                      Date : \n Amount :                                        Risk :", PolicyNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 3;
                    cell.PaddingLeft = 10;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(CompanyGSTIN, VRCGSTIN));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 8;
                    cell.Padding = 5;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Consignor Name :-", txtConName));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsignorName"].ToString(), ConName));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    if (dt.Rows[0]["MType"].ToString() == "B")
                    {
                        cell = new PdfPCell(new Phrase("Consignment No (" + (z + 1) + "/" + dt.Rows.Count.ToString() + ") :", txtConNo));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthBottom = 1;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 1;
                        table.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase("Consignment No :", txtConNo));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 0;
                        cell.BorderWidthBottom = 1;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 1;
                        table.AddCell(cell);
                    }

                    cell = new PdfPCell(new Phrase("Address : -", ConAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsignorAdd"].ToString(), ConAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookLrNo"].ToString(), ConNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Date :- " + dt.Rows[z]["BookDate1"].ToString(), Date));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsignorState"].ToString(), ConAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("GST NO :-", ConGSTIN));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsignorGST"].ToString(), ConGSTIN));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Vehicle No:-", txtVehicleNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Consignee Name", txtConName));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["Consigneename"].ToString(), ConName));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["VchNo"].ToString(), VechicleNo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Address : -", ConAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsigneeAdd"].ToString() + "," + dt.Rows[z]["ConsigneeState"].ToString(), ConAddress));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("From :- " + dt.Rows[z]["SourceFrom"].ToString(), FromTo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("To :- " + dt.Rows[z]["ToDestination"].ToString(), FromTo));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("GST NO : - ", ConGSTIN));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    cell.Padding = 4;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ConsigneeGST"].ToString(), ConGSTIN));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Packages", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Description(said to contain)", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    DataTable dtLock = con.GetDatatable("select * from [tblVajraLock] where vehicleno='" + dt.Rows[z]["VchNo"].ToString() + "'");
                    if (dtLock.Rows.Count > 0)
                    {
                        cell = new PdfPCell(new Phrase("VAJRA LOCK - INSTALLED", notice));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 0;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase("VAJRA LOCK - ", notice));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 0;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }


                    cell = new PdfPCell(new Phrase("Rate", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Amount", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Shipment No:- " + dt.Rows[z]["BookShipNo"].ToString(), fontB));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookQty"].ToString() + "," + dt.Rows[z]["UnitName"].ToString(), font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 4;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["ProdName"].ToString(), fontB));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Rowspan = 4;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Vehicle Size", VechicleSize));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Frieght", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("LC No :- " + dt.Rows[z]["LCNo"].ToString(), fontB));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookCustSize"].ToString(), VechicleSize));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Hamali", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Doc. No:-" + DocumentNo, fontB));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_TOP;
                    cell.Rowspan = 3;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Collection Charges", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 2;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Total:-  " + dt.Rows[z]["BookQty"].ToString() + "," + dt.Rows[z]["UnitName"].ToString(), font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Seal No:- " + dt.Rows[z]["SealNo"].ToString(), fontB));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 2;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    if (dt.Rows[z]["BookBillTo"].ToString() == "TOPAY")
                    {
                        cell = new PdfPCell(new Phrase("To be billed To :- " + dt.Rows[z]["Consigneename"].ToString(), font));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }
                    else
                    {
                        cell = new PdfPCell(new Phrase("To be billed To :- " + dt.Rows[z]["BookBillTo"].ToString(), font));
                        cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                        cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                        cell.Padding = 4;
                        cell.BorderWidthTop = 1;
                        cell.BorderWidthBottom = 0;
                        cell.BorderWidthLeft = 1;
                        cell.BorderWidthRight = 0;
                        table.AddCell(cell);
                    }

                    cell = new PdfPCell(new Phrase("Total", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);





                    cell = new PdfPCell(new Phrase(InvoiceDetails, font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 7;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);



                    cell = new PdfPCell(new Phrase("Barcode", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    cell.Padding = 4;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(EwayDetails, font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 7;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    imageBytes = Convert.FromBase64String(StrBase64);
                    jpg = iTextSharp.text.Image.GetInstance(imageBytes);
                    jpg.ScaleAbsolute(130f, 29.0f);
                    cell = new PdfPCell(jpg);
                    cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Rowspan = 2;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Remark", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookRemark1"].ToString(), font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 6;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 0;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Goods accepted for carriage of the terms & conditions Printed overleaf.", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Colspan = 5;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase("Signature of Transport Operator", Sign));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 3;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(" ", font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 0;
                    cell.BorderWidthTop = 1;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 0;
                    cell.BorderWidthRight = 0;
                    table.AddCell(cell);

                    cell = new PdfPCell(new Phrase(dt.Rows[z]["BookLoading"].ToString(), font));
                    cell.HorizontalAlignment = PdfPCell.ALIGN_LEFT;
                    cell.VerticalAlignment = PdfPCell.ALIGN_CENTER;
                    cell.Padding = 4;
                    cell.BorderWidthTop = 0;
                    cell.BorderWidthBottom = 1;
                    cell.BorderWidthLeft = 1;
                    cell.BorderWidthRight = 1;
                    table.AddCell(cell);

                    document.Add(table);

                    document.NewPage();
                }


            }

            document.Close();
            document.Dispose();
            writer.Close();
            writer.Dispose();
            fs.Close();
            fs.Dispose();
        }
        private string FolderCheck(string FolderName)
        {
            string ServerPath = string.Empty;
            ServerPath = System.Web.HttpContext.Current.Server.MapPath("~/" + FolderName + "/");
            if (!Directory.Exists(ServerPath))
                Directory.CreateDirectory(ServerPath);
            return ServerPath;
        }
        /*
        private string GenerateBarCode(String LrNo)
        {
            LrNo = LrNo.Trim();
            string ImagePath = "/App_Content/BarcodeImage/BrCode_" + LrNo + ".jpg";
            string barcodePath = System.Web.Hosting.HostingEnvironment.MapPath(ImagePath);
            // Generate a Simple BarCode image and save as PNG
            IronBarCode.GeneratedBarcode BarCode = IronBarCode.BarcodeWriter.CreateBarcode("https://ironsoftware.com/csharp/barcode", BarcodeWriterEncoding.Code128);
            BarCode.SaveAsPng(barcodePath);
            // This line opens the image in your default image viewerSystem.Diagnostics.Process.Start("BarCode.png");

            //Converting Image to Base64
            string base64String = string.Empty;
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(barcodePath))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();

                    // Convert byte[] to Base64 String
                    base64String = Convert.ToBase64String(imageBytes);
                    //return base64String;
                }
            }
            File.Delete(barcodePath);
            return base64String;
        }*/

        private string GenerateBarCode(String LrNo)
        {
            LrNo = LrNo.Trim();
            int pos = LrNo.LastIndexOf("/") + 1;
            LrNo = LrNo.Substring(pos, LrNo.Length - pos);

            var folderPath = System.Web.Hosting.HostingEnvironment.MapPath("~") + "/App_Content/BarcodeImage";
            string ImagePath = "/App_Content/BarcodeImage/BrCode_" + LrNo + ".jpg";
            var barcodeWriter = new BarcodeWriter();
            barcodeWriter.Format = BarcodeFormat.CODABAR;
            var result = barcodeWriter.Write(LrNo);
            string barcodePath = System.Web.Hosting.HostingEnvironment.MapPath(ImagePath);
            var barcodeBitmap = new Bitmap(result);
            using (MemoryStream memory = new MemoryStream())
            {
                using (FileStream fs = new FileStream(barcodePath, FileMode.Create, FileAccess.ReadWrite))
                {
                    barcodeBitmap.Save(memory, ImageFormat.Jpeg);
                    byte[] bytes = memory.ToArray();
                    fs.Write(bytes, 0, bytes.Length);
                }
            }

            //Converting Image to Base64
            string base64String = string.Empty;
            using (System.Drawing.Image image = System.Drawing.Image.FromFile(barcodePath))
            {
                using (MemoryStream m = new MemoryStream())
                {
                    image.Save(m, image.RawFormat);
                    byte[] imageBytes = m.ToArray();

                    // Convert byte[] to Base64 String
                    base64String = Convert.ToBase64String(imageBytes);
                    //return base64String;
                }
            }
            File.Delete(barcodePath);
            return base64String;
        }
        #endregion

        #region[Events]
        protected void btnConsignor_Click(object sender, EventArgs e)
        {
            try
            {
                if (Request.QueryString["BookId"] != null)
                {
                    string strHTML = string.Empty;
                    string Reporttxt = "CONSIGNOR COPY";
                    string FileName = "Invoice_Bid" + Request.QueryString["BookId"] + "_" + DateTime.Now.ToString("HH-mm-ss.fff") + ".pdf";
                    string FinalURL = System.Configuration.ConfigurationManager.AppSettings["ConfigPath"].ToString() + "App_Content/Report/" + FileName;
                    //PrintAllReport_iText(Request.QueryString["BookId"], FileName, Reporttxt);
                    GenerateLRPdf(Request.QueryString["BookId"], FileName, Reporttxt);
                    Response.Redirect(FinalURL, false);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());

            }
        }
        protected void btnConsignee_Click(object sender, EventArgs e)
        {
            try
            {
                if (Request.QueryString["BookId"] != null)
                {
                    string strHTML = string.Empty;
                    string Reporttxt = "CONSIGNEE COPY";
                    string FileName = "Invoice_Bid" + Request.QueryString["BookId"] + "_" + DateTime.Now.ToString("HH-mm-ss.fff") + ".pdf";
                    string FinalURL = System.Configuration.ConfigurationManager.AppSettings["ConfigPath"].ToString() + "App_Content/Report/" + FileName;
                    GenerateLRPdf(Request.QueryString["BookId"], FileName, Reporttxt);
                    Response.Redirect(FinalURL, false);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());

            }
        }
        protected void btnCustomer_Click(object sender, EventArgs e)
        {
            try
            {
                if (Request.QueryString["BookId"] != null)
                {
                    string strHTML = string.Empty;
                    string Reporttxt = "ACCOUNT COPY";
                    string FileName = "Invoice_Bid" + Request.QueryString["BookId"] + "_" + DateTime.Now.ToString("HH-mm-ss.fff") + ".pdf";
                    string FinalURL = System.Configuration.ConfigurationManager.AppSettings["ConfigPath"].ToString() + "App_Content/Report/" + FileName;
                    GenerateLRPdf(Request.QueryString["BookId"], FileName, Reporttxt);
                    Response.Redirect(FinalURL, false);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());

            }
        }
        protected void btnDriver_Click(object sender, EventArgs e)
        {
            try
            {
                if (Request.QueryString["BookId"] != null)
                {
                    string strHTML = string.Empty;
                    string Reporttxt = "DRIVER COPY";
                    string FileName = "Invoice_Bid" + Request.QueryString["BookId"] + "_" + DateTime.Now.ToString("HH-mm-ss.fff") + ".pdf";
                    string FinalURL = System.Configuration.ConfigurationManager.AppSettings["ConfigPath"].ToString() + "App_Content/Report/" + FileName;
                    GenerateLRPdf(Request.QueryString["BookId"], FileName, Reporttxt);
                    Response.Redirect(FinalURL, false);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());

            }
        }
        protected void btnAll_Click(object sender, EventArgs e)
        {
            try
            {
                if (Request.QueryString["BookId"] != null)
                {
                    string strHTML = string.Empty;
                    string Reporttxt = "";
                    string FileName = "Invoice_Bid" + Request.QueryString["BookId"] + "_" + DateTime.Now.ToString("HH-mm-ss.fff") + ".pdf";
                    string FinalURL = System.Configuration.ConfigurationManager.AppSettings["ConfigPath"].ToString() + "App_Content/Report/" + FileName;
                    GenerateLRPdf(Request.QueryString["BookId"], FileName, Reporttxt);
                    Response.Redirect(FinalURL, false);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());

            }
        }
        #endregion

    }
}
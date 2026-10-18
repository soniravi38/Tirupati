using Consignment.Model;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Org.BouncyCastle.Asn1.X509;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Consignment.ConsignmentPanel
{
    public partial class ConsgReport : System.Web.UI.Page
    {
        #region [properties]
        ArrayList ParameterName, ParameterValue;
        DataTable dt = new DataTable();
        DataSet ds = new DataSet();
        connect con = new connect();
        common cm = new common();
        #endregion

        #region [pageload]
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                fillList();

                GridViewSortDirection = SortDirection.Descending;
                lblHeading.Text = "Consignment Report";
            }
        }
        #endregion

        #region [All-Function]
        private void fillList()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "[Proc_ConsignmentReportLR]";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("ViewConsignment");
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ParameterName.Add("@BookBranchId"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["BranchId"]).ToString());

                ParameterName.Add("@FromDate"); ParameterValue.Add(txtDateFrom.Text);
                ParameterName.Add("@ToDate"); ParameterValue.Add(txtDateTo.Text);
                ParameterName.Add("@LrFrom"); ParameterValue.Add(txtLRFrom.Text);
                ParameterName.Add("@LrTo"); ParameterValue.Add(txtLRTo.Text);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                if (ds.Tables.Count > 0)
                {
                    dt = ds.Tables[0];
                    gvReport.PageSize = Convert.ToInt32(ddlPageSize.SelectedValue);
                    gvReport.DataSource = dt;
                    gvReport.DataBind();
                    litrecordcount.Text = dt.Rows.Count.ToString();
                    ViewState["dtExport"] = dt;
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }

        }
        public SortDirection GridViewSortDirection
        {
            get
            {
                if (ViewState["sortDirection"] == null)
                {
                    ViewState["sortDirection"] = SortDirection.Ascending;
                }
                return (SortDirection)ViewState["sortDirection"];
            }
            set
            {
                ViewState["sortDirection"] = value;
            }
        }
        #endregion

        #region [GridViewEvents]
      
        protected void gvReport_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            gvReport.PageIndex = e.NewPageIndex;
            fillList();
        }
        protected void gvReport_Sorting(object sender, GridViewSortEventArgs e)
        {
            try
            {
                DataTable dt = (DataTable)ViewState["dtExport"];
                DataView dv = new DataView(dt);
                if (GridViewSortDirection == SortDirection.Ascending)
                {
                    GridViewSortDirection = SortDirection.Descending;
                    dv.Sort = e.SortExpression + " DESC";
                }
                else
                {
                    GridViewSortDirection = SortDirection.Ascending;
                    dv.Sort = e.SortExpression + " ASC";
                }
                gvReport.DataSource = dv;
                gvReport.DataBind();
            }
            catch (Exception ex)
            { cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString()); }
        }
        #endregion

        #region [ddlPaging]
        protected void ddlPageSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            fillList();
        }
        #endregion

        #region[Events]
        protected void btnSearch_Click(object sender, EventArgs e)
        {
            fillList();
        }

        protected void imgBtnExcel_Click(object sender, ImageClickEventArgs e)
        {
            DataTable dtTemp = (DataTable)ViewState["dtExport"];
            if (dtTemp.Rows.Count > 0)
            {

                cm.ExportToExcel(dtTemp, "LR");
            }
        }



        protected void imgBtnPdf_Click(object sender, ImageClickEventArgs e)
        {
            DataTable dtTemp = (DataTable)ViewState["dtExport"];
            if (dtTemp.Rows.Count > 0)
            {

                ExportToPdf(dtTemp, "LR");
            }
        }

        public void ExportToPdf(DataTable dt, string fileName)
        {
            try
            {
                if (dt == null || dt.Rows.Count == 0)
                    return;

                string folderPath = HttpContext.Current.Server.MapPath("~/App_Content/Report/");

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                string fullPath = Path.Combine(
                    folderPath,
                    fileName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf"
                );

                // A3 Landscape because there are many columns
                Document document = new Document(
                    PageSize.A3.Rotate(),
                    20f,
                    20f,
                    25f,
                    25f
                );

                PdfWriter writer = PdfWriter.GetInstance(
                    document,
                    new FileStream(fullPath, FileMode.Create)
                );

                document.Open();

                // =========================================================
                // FONTS
                // =========================================================

                Font titleFont = FontFactory.GetFont(
                    "Arial",
                    15,
                    Font.BOLD,
                    new BaseColor(31, 78, 121)
                );

                Font headerFont = FontFactory.GetFont(
                    "Arial",
                    7,
                    Font.BOLD,
                    BaseColor.WHITE
                );

                Font bodyFont = FontFactory.GetFont(
                    "Arial",
                    6.5f,
                    Font.NORMAL,
                    new BaseColor(40, 40, 40)
                );

                Font smallFont = FontFactory.GetFont(
                    "Arial",
                    6,
                    Font.NORMAL,
                    new BaseColor(60, 60, 60)
                );

                // =========================================================
                // TITLE
                // =========================================================

                Paragraph title = new Paragraph(
                    "LR BOOK REPORT",
                    titleFont
                );

                title.Alignment = Element.ALIGN_CENTER;
                title.SpacingAfter = 8f;

                document.Add(title);

                // =========================================================
                // REPORT DATE
                // =========================================================

                Paragraph reportDate = new Paragraph(
                    "Generated On : " + DateTime.Now.ToString("dd-MM-yyyy HH:mm"),
                    smallFont
                );

                reportDate.Alignment = Element.ALIGN_RIGHT;
                reportDate.SpacingAfter = 8f;

                document.Add(reportDate);

                // =========================================================
                // TABLE
                // =========================================================

                int columnCount = dt.Columns.Count;

                PdfPTable table = new PdfPTable(columnCount);

                table.WidthPercentage = 100;

                table.HeaderRows = 1;

                table.SplitLate = false;
                table.SplitRows = true;

                // =========================================================
                // COLUMN WIDTHS
                // =========================================================

                float[] widths = new float[columnCount];

                for (int i = 0; i < columnCount; i++)
                {
                    string columnName = dt.Columns[i].ColumnName;

                    switch (columnName)
                    {
                        case "Bookid":
                            widths[i] = 4;
                            break;

                        case "BookLrNo":
                            widths[i] = 6;
                            break;

                        case "BookDate1":
                            widths[i] = 7;
                            break;

                        case "BranchName":
                        case "SourceFrom":
                        case "ToDestination":
                            widths[i] = 8;
                            break;

                        case "CreatedBy":
                        case "UpdatedBy":
                            widths[i] = 7;
                            break;

                        case "Created_Date":
                        case "Modified_Date":
                            widths[i] = 9;
                            break;

                        case "ConsignorName":
                        case "Consigneename":
                            widths[i] = 13;
                            break;

                        case "ConsignorAdd":
                        case "ConsigneeAdd":
                            widths[i] = 18;
                            break;

                        case "ConsignorState":
                        case "ConsigneeState":
                            widths[i] = 8;
                            break;

                        case "ConsignorGST":
                        case "ConsigneeGST":
                            widths[i] = 10;
                            break;

                        case "VchNo":
                            widths[i] = 8;
                            break;

                        case "VchSize":
                            widths[i] = 6;
                            break;

                        case "BookBillTo1":
                            widths[i] = 12;
                            break;

                        case "DocumentDetails":
                        case "InvoiceDetails":
                        case "EwayBillDetails":
                            widths[i] = 18;
                            break;

                        default:
                            widths[i] = 8;
                            break;
                    }
                }

                table.SetWidths(widths);

                // =========================================================
                // HEADER
                // =========================================================

                foreach (DataColumn column in dt.Columns)
                {
                    string headerText = column.ColumnName;

                    // Make headers user friendly
                    switch (headerText)
                    {
                        case "Bookid":
                            headerText = "Book ID";
                            break;

                        case "BookLrNo":
                            headerText = "LR No";
                            break;

                        case "BookDate1":
                            headerText = "LR Date";
                            break;

                        case "SourceFrom":
                            headerText = "From";
                            break;

                        case "ToDestination":
                            headerText = "To";
                            break;

                        case "Consigneename":
                            headerText = "Consignee Name";
                            break;

                        case "ConsignorName":
                            headerText = "Consignor Name";
                            break;

                        case "ConsignorAdd":
                            headerText = "Consignor Address";
                            break;

                        case "ConsigneeAdd":
                            headerText = "Consignee Address";
                            break;

                        case "ConsignorGST":
                            headerText = "Consignor GST";
                            break;

                        case "ConsigneeGST":
                            headerText = "Consignee GST";
                            break;

                        case "VchNo":
                            headerText = "Vehicle No";
                            break;

                        case "VchSize":
                            headerText = "Vehicle Size";
                            break;

                        case "BookBillTo1":
                            headerText = "Bill To";
                            break;

                        case "DocumentDetails":
                            headerText = "Document Details";
                            break;

                        case "InvoiceDetails":
                            headerText = "Invoice Details";
                            break;

                        case "EwayBillDetails":
                            headerText = "E-Way Bill Details";
                            break;

                        case "Created_Date":
                            headerText = "Created Date";
                            break;

                        case "Modified_Date":
                            headerText = "Modified Date";
                            break;
                    }

                    PdfPCell headerCell = new PdfPCell(
                        new Phrase(headerText, headerFont)
                    );

                    // Professional dark blue header
                    headerCell.BackgroundColor = new BaseColor(31, 78, 121);

                    headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    headerCell.VerticalAlignment = Element.ALIGN_MIDDLE;

                    headerCell.PaddingTop = 5f;
                    headerCell.PaddingBottom = 5f;

                    headerCell.BorderColor = new BaseColor(255, 255, 255);
                    headerCell.BorderWidth = 0.5f;

                    table.AddCell(headerCell);
                }

                // =========================================================
                // DATA ROWS
                // =========================================================

                for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
                {
                    DataRow row = dt.Rows[rowIndex];

                    // Alternate row colors
                    BaseColor rowColor;

                    if (rowIndex % 2 == 0)
                    {
                        rowColor = new BaseColor(245, 248, 252);
                    }
                    else
                    {
                        rowColor = BaseColor.WHITE;
                    }

                    foreach (DataColumn column in dt.Columns)
                    {
                        string value = "";

                        if (row[column] != DBNull.Value)
                            value = row[column].ToString();

                        PdfPCell dataCell = new PdfPCell(
                            new Phrase(value, bodyFont)
                        );

                        dataCell.BackgroundColor = rowColor;

                        dataCell.VerticalAlignment = Element.ALIGN_TOP;

                        // Long text should wrap
                        dataCell.NoWrap = false;

                        dataCell.PaddingTop = 3f;
                        dataCell.PaddingBottom = 3f;
                        dataCell.PaddingLeft = 3f;
                        dataCell.PaddingRight = 3f;

                        dataCell.BorderColor = new BaseColor(190, 200, 210);
                        dataCell.BorderWidth = 0.4f;

                        // Center short fields
                        string colName = column.ColumnName;

                        if (colName == "Bookid" ||
                            colName == "BookLrNo" ||
                            colName == "BookDate1" ||
                            colName == "BranchName" ||
                            colName == "SourceFrom" ||
                            colName == "ToDestination" ||
                            colName == "CreatedBy" ||
                            colName == "UpdatedBy" ||
                            colName == "VchNo" ||
                            colName == "VchSize")
                        {
                            dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                        }
                        else
                        {
                            dataCell.HorizontalAlignment = Element.ALIGN_LEFT;
                        }

                        table.AddCell(dataCell);
                    }
                }

                // =========================================================
                // ADD TABLE
                // =========================================================

                document.Add(table);

                // =========================================================
                // FOOTER
                // =========================================================

                Paragraph footer = new Paragraph(
                    "Total Records : " + dt.Rows.Count,
                    smallFont
                );

                footer.Alignment = Element.ALIGN_RIGHT;
                footer.SpacingBefore = 6f;

                document.Add(footer);

                document.Close();
                writer.Close();

                // =========================================================
                // DOWNLOAD
                // =========================================================

                HttpResponse response = HttpContext.Current.Response;

                response.Clear();
                response.ClearHeaders();
                response.ClearContent();

                response.ContentType = "application/pdf";
                response.AddHeader(
                    "Content-Disposition",
                    "attachment; filename=" +
                    fileName + ".pdf"
                );

                response.TransmitFile(fullPath);
                response.Flush();

                HttpContext.Current.ApplicationInstance.CompleteRequest();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using Consignment.Model;
using System.Collections;
using static System.Net.WebRequestMethods;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;

namespace Consignment.ConsignmentPanel
{
    public partial class AddBookLR : System.Web.UI.Page
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
            // Session expire ho chuka ho to crash (yellow error page) ki jagah seedha login par bhejo
            if (Session["UserId"] == null || Session["UserName"] == null || Session["Database_"] == null)
            {
                Response.Redirect("Default.aspx", false);
                Context.ApplicationInstance.CompleteRequest();
                return;
            }
            divsave.Visible = false;
            hfbrid.Value = (cm.GetUserDetails().Rows[0]["BranchId"]).ToString();
            if(Session["Database_"].ToString() == "tirupati_")
            {
                lblSize.Text = "Size as per Tirupati";
            }
            else if(Session["Database_"].ToString() == "hpexpress_")
            {
                lblSize.Text = "Size as per HP";
            }
            else
            {
                lblSize.Text = "Size as per Company";
                
            }
            lblSize.Font.Bold = true;
            FillProducts();
            if (!IsPostBack)
            {
                FillConsignor();
                FillConsignee();
                FillBillingParty();
                FillSource();
                FillUnit();
                FillSize();
                Clear();
                txtdate.Text = cm.GetCurrentDate();
                if (Request.QueryString["BookId"] != null)
                {
                    HFid.Value = Request.QueryString["BookId"].ToString();
                    GetLRData(Convert.ToInt32(HFid.Value));
                    ddConsignee_SelectedIndexChanged(sender, e);
                    ddConsignor_SelectedIndexChanged(sender, e);
                    ddlVchNo_SelectedIndexChanged(sender, e);
                }
            }
        }
        #endregion

        #region [all-fuctions]
        private void FillConsignor()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BranchwiseConsig";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetPartyByBranchId");
                ParameterName.Add("@BranchId"); ParameterValue.Add(Convert.ToInt32(hfbrid.Value));
                ParameterName.Add("@ConsgMode"); ParameterValue.Add("A");
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddConsignor.DataSource = dt;
                    ddConsignor.DataValueField = "Id";
                    ddConsignor.DataTextField = "Consgname";
                    ddConsignor.DataBind();
                    ddConsignor.Items.Insert(0, new ListItem("--Select Consignor--", "0"));
                }
                else
                {
                    ddConsignor.Items.Clear();
                    ddConsignor.Items.Insert(0, new ListItem("--Select Consignor--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        private void FillConsignee()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BranchwiseConsig";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetPartyByBranchId");
                ParameterName.Add("@BranchId"); ParameterValue.Add(Convert.ToInt32(hfbrid.Value));
                ParameterName.Add("@ConsgMode"); ParameterValue.Add("S");
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddConsignee.DataSource = dt;
                    ddConsignee.DataValueField = "Id";
                    ddConsignee.DataTextField = "Consgname";
                    ddConsignee.DataBind();
                    ddConsignee.Items.Insert(0, new ListItem("--Select Consignee--", "0"));
                }
                else
                {
                    ddConsignee.Items.Clear();
                    ddConsignee.Items.Insert(0, new ListItem("--Select Consignee--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        private void FillBillingParty()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BranchwiseConsig";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetPartyByBranchId");
                ParameterName.Add("@BranchId"); ParameterValue.Add(Convert.ToInt32(hfbrid.Value));
                ParameterName.Add("@ConsgMode"); ParameterValue.Add("B");
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddlBillto.DataSource = dt;
                    ddlBillto.DataValueField = "Consgname";
                    ddlBillto.DataTextField = "Consgname";
                    ddlBillto.DataBind();
                    ddlBillto.Items.Insert(0, new ListItem("--Select Billing Party--", "0"));
                }
                else
                {
                    ddlBillto.Items.Clear();
                    ddlBillto.Items.Insert(0, new ListItem("--Select Billing Party--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        private void FillSource()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_Station";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("List");
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddSourceFrom.DataSource = dt;
                    ddSourceFrom.DataValueField = "Id";
                    ddSourceFrom.DataTextField = "Name";
                    ddSourceFrom.DataBind();
                    ddSourceFrom.Items.Insert(0, new ListItem("--Select Source From--", "0"));

                    ddSourceTo.DataSource = dt;
                    ddSourceTo.DataValueField = "Id";
                    ddSourceTo.DataTextField = "Name";
                    ddSourceTo.DataBind();
                    ddSourceTo.Items.Insert(0, new ListItem("--Select Source To--", "0"));
                }
                else
                {
                    ddSourceFrom.Items.Clear();
                    ddSourceFrom.Items.Insert(0, new ListItem("--Select Source From--", "0"));

                    ddSourceTo.Items.Clear();
                    ddSourceTo.Items.Insert(0, new ListItem("--Select Source To--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        private void FillVehicleById(int VehicleId)
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_Vehicle";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("Get");
                ParameterName.Add("@Id"); ParameterValue.Add(VehicleId);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                DataTable dt1 = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddlVchNo.DataSource = dt1;
                    ddlVchNo.DataValueField = "Id";
                    ddlVchNo.DataTextField = "VehicleNo";
                    ddlVchNo.DataBind();
                }
                else
                {
                    ddlVchNo.Items.Clear();
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        private void FillSize()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_Size";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("List");
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddlSize.DataSource = dt;
                    ddlSize.DataValueField = "Name";
                    ddlSize.DataTextField = "Name";
                    ddlSize.DataBind();
                    ddlSize.Items.Insert(0, new ListItem("--Select Size--", "0"));
                }
                else
                {
                    ddlSize.Items.Clear();
                    ddlSize.Items.Insert(0, new ListItem("--Select Size--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        // Product Description dropdown (datalist): tblProduct master se list
        private void FillProducts()
        {
            string spName = "Proc_Product";
            try
            {
                ArrayList pn = new ArrayList(); ArrayList pv = new ArrayList();
                pn.Add("@Action"); pv.Add("List");
                DataSet pds = con.GetDataSetFromSP(spName, pn, pv);
                StringBuilder sb = new StringBuilder();
                if (pds.Tables.Count > 0)
                {
                    foreach (DataRow r in pds.Tables[0].Rows)
                    {
                        sb.Append("<option value=\"").Append(HttpUtility.HtmlAttributeEncode(r["ProductName"].ToString())).Append("\"></option>");
                    }
                }
                litProducts.Text = sb.ToString();
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:FillProducts | SP:" + spName + " | User:" + Convert.ToString(Session["UserName"]) + " | Err:" + ex.Message.ToString());
            }
        }

        // Naya product type kiya ho to master me jod do (pehle se ho to kuch nahi hota)
        private void SaveProduct(string productName)
        {
            if (string.IsNullOrWhiteSpace(productName)) return;
            string spName = "Proc_Product";
            try
            {
                ArrayList pn = new ArrayList(); ArrayList pv = new ArrayList();
                pn.Add("@Action"); pv.Add("AddIfNotExists");
                pn.Add("@ProductName"); pv.Add(productName.Trim());
                pn.Add("@UserId"); pv.Add(Convert.ToInt32(Session["UserId"]));
                con.ExecuteQueryByStoreProcedureWithParameter(spName, pn, pv);
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:SaveProduct | SP:" + spName + " | User:" + Convert.ToString(Session["UserName"]) + " | Err:" + ex.Message.ToString());
            }
        }

        private void FillUnit()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_Unit";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("List");
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddlunit.DataSource = dt;
                    ddlunit.DataValueField = "Name";
                    ddlunit.DataTextField = "Name";
                    ddlunit.DataBind();
                    ddlunit.Items.Insert(0, new ListItem("--Select Unit--", "0"));
                }
                else
                {
                    ddlunit.Items.Clear();
                    ddlunit.Items.Insert(0, new ListItem("--Select Unit--", "0"));
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        public void Clear()
        {
            txtLCNo.Text = "";

            txtAmount.Text = "0";
            txtdate.Text = "";
            txtDestination.Text = "";
            txtDocNo.Text = "";
            txtEwayBillNo.Text = "";
            txtfromdate.Text = "";
            txtinvdate.Text = "";
            txtinvno.Text = "";
            txtLoading.Text = "";
            TxtLrNo.Text = "";
            txtQty.Text = "0";
            TxtLrNo.ReadOnly = false;
            txtRemarks.Text = "";
            txtShipNo.Text = "";
            txttodate.Text = "";
            HFid.Value = "";
            lblConsigneeAddress.Text = "";
            lblConsigneeGst.Text = "";
            lblConsigneeState.Text = "";
            lblConsignorAddress.Text = "";
            lblConsignorGST.Text = "";
            lblConsignorState.Text = "";
            ddlBillto.SelectedIndex = 0;
            txtsealno.Text = "";
            txtDriverMobile.Text = "";
            txtDriverAdhar.Text = "";
            ddlSize.SelectedIndex = 0;
            txtRemarks1.Text = "";
            ddConsignee.SelectedIndex = 0;
            txtProduct.Text = "";
            ddlGSTPay.SelectedIndex = 0;
            ddlunit.SelectedIndex = 0;
            ddlVchNo.SelectedIndex = 0;
            ddSourceFrom.SelectedIndex = 0;
            ddSourceTo.SelectedIndex = 0;
            ddVType.SelectedIndex = 0;
            TxtLrNo.Text = "";
            if (GetLastLrNo() != "")
            {
                lblcongo.Text = GetLastLrNo();
            }

            txtDocumentNo1.Text = "";
            txtDocumentNo2.Text = "";
            txtDocumentNo3.Text = "";
            txtDocumentNo4.Text = "";
            txtDocumentNo5.Text = "";
            txtDocumentNo6.Text = "";
            txtDocumentNo7.Text = "";
            txtDocumentNo8.Text = "";
            txtDocumentNo9.Text = "";
            txtDocumentNo10.Text = "";
            txtDocumentNo11.Text = "";
            txtDocumentNo12.Text = "";
            txtDocumentNo13.Text = "";
            txtDocumentNo14.Text = "";
            txtDocumentNo15.Text = "";

            txtInvoiceNo1.Text = "";
            InvoiceDate1.Text = "";
            InvoiceValue1.Text = "";
            InvoiceUnit1.Text = "";

            txtInvoiceNo2.Text = "";
            InvoiceDate2.Text = "";
            InvoiceValue2.Text = "";
            InvoiceUnit2.Text = "";

            txtInvoiceNo3.Text = "";
            InvoiceDate3.Text = "";
            InvoiceValue3.Text = "";
            InvoiceUnit3.Text = "";

            txtInvoiceNo4.Text = "";
            InvoiceDate4.Text = "";
            InvoiceValue4.Text = "";
            InvoiceUnit4.Text = "";

            txtInvoiceNo5.Text = "";
            InvoiceDate5.Text = "";
            InvoiceValue5.Text = "";
            InvoiceUnit5.Text = "";

            txtInvoiceNo6.Text = "";
            InvoiceDate6.Text = "";
            InvoiceValue6.Text = "";
            InvoiceUnit6.Text = "";

            txtInvoiceNo7.Text = "";
            InvoiceDate7.Text = "";
            InvoiceValue7.Text = "";
            InvoiceUnit7.Text = "";

            txtInvoiceNo8.Text = "";
            InvoiceDate8.Text = "";
            InvoiceValue8.Text = "";
            InvoiceUnit8.Text = "";

            txtInvoiceNo9.Text = "";
            InvoiceDate9.Text = "";
            InvoiceValue9.Text = "";
            InvoiceUnit9.Text = "";

            txtInvoiceNo10.Text = "";
            InvoiceDate10.Text = "";
            InvoiceValue10.Text = "";
            InvoiceUnit10.Text = "";

            txtInvoiceNo11.Text = "";
            InvoiceDate11.Text = "";
            InvoiceValue11.Text = "";
            InvoiceUnit11.Text = "";

            txtInvoiceNo12.Text = "";
            InvoiceDate12.Text = "";
            InvoiceValue12.Text = "";
            InvoiceUnit12.Text = "";

            txtInvoiceNo13.Text = "";
            InvoiceDate13.Text = "";
            InvoiceValue13.Text = "";
            InvoiceUnit13.Text = "";

            txtInvoiceNo14.Text = "";
            InvoiceDate14.Text = "";
            InvoiceValue14.Text = "";
            InvoiceUnit14.Text = "";

            txtInvoiceNo15.Text = "";
            InvoiceDate15.Text = "";
            InvoiceValue15.Text = "";
            InvoiceUnit15.Text = "";

            txtEwayBillNo1.Text = "";
            txtEwayBillDate1.Text = "";
            txtValidUptoDate1.Text = "";


            txtEwayBillNo2.Text = "";
            txtEwayBillDate2.Text = "";
            txtValidUptoDate2.Text = "";

            txtEwayBillNo3.Text = "";
            txtEwayBillDate3.Text = "";
            txtValidUptoDate3.Text = "";

            txtEwayBillNo4.Text = "";
            txtEwayBillDate4.Text = "";
            txtValidUptoDate4.Text = "";

            txtEwayBillNo5.Text = "";
            txtEwayBillDate5.Text = "";
            txtValidUptoDate5.Text = "";

            txtEwayBillNo6.Text = "";
            txtEwayBillDate6.Text = "";
            txtValidUptoDate6.Text = "";

            txtEwayBillNo7.Text = "";
            txtEwayBillDate7.Text = "";
            txtValidUptoDate7.Text = "";

            txtEwayBillNo8.Text = "";
            txtEwayBillDate8.Text = "";
            txtValidUptoDate8.Text = "";

            txtEwayBillNo9.Text = "";
            txtEwayBillDate9.Text = "";
            txtValidUptoDate9.Text = "";

            txtEwayBillNo10.Text = "";
            txtEwayBillDate10.Text = "";
            txtValidUptoDate10.Text = "";

            txtEwayBillNo11.Text = "";
            txtEwayBillDate11.Text = "";
            txtValidUptoDate11.Text = "";

            txtEwayBillNo12.Text = "";
            txtEwayBillDate12.Text = "";
            txtValidUptoDate12.Text = "";

            txtEwayBillNo13.Text = "";
            txtEwayBillDate13.Text = "";
            txtValidUptoDate13.Text = "";

            txtEwayBillNo14.Text = "";
            txtEwayBillDate14.Text = "";
            txtValidUptoDate14.Text = "";

            txtEwayBillNo15.Text = "";
            txtEwayBillDate15.Text = "";
            txtValidUptoDate15.Text = "";
        }
        public void ClearForMulti()
        {
            //txtLCNo.Text = "";
            txtAmount.Text = "0";
            txtdate.Text = "";
            txtDestination.Text = "";
            txtDocNo.Text = "";
            txtEwayBillNo.Text = "";
            txtfromdate.Text = "";
            txtinvdate.Text = "";
            txtinvno.Text = "";
            txtLoading.Text = "";
            TxtLrNo.Text = "";
            txtQty.Text = "0";
            TxtLrNo.ReadOnly = false;
            txtRemarks.Text = "";
            txtShipNo.Text = "";
            txttodate.Text = "";
            HFid.Value = "";
            lblConsigneeAddress.Text = "";
            lblConsigneeGst.Text = "";
            lblConsigneeState.Text = "";
            ddlGSTPay.SelectedIndex = 0;
            ddlBillto.SelectedIndex = 0;
            txtsealno.Text = "";
            txtDriverMobile.Text = "";
            txtDriverAdhar.Text = "";
            txtRemarks1.Text = "";
            ddConsignee.SelectedIndex = 0;
            txtProduct.Text = "";
            ddlunit.SelectedIndex = 0;
            ddSourceTo.SelectedIndex = 0;
            ddVType.SelectedIndex = 0;
            TxtLrNo.Text = "";
            if (GetLastLrNo() != "")
            {
                lblcongo.Text = GetLastLrNo();
            }
            txtdate.Text = cm.GetCurrentDate();

            txtDocumentNo1.Text = "";
            txtDocumentNo2.Text = "";
            txtDocumentNo3.Text = "";
            txtDocumentNo4.Text = "";
            txtDocumentNo5.Text = "";
            txtDocumentNo6.Text = "";
            txtDocumentNo7.Text = "";
            txtDocumentNo8.Text = "";
            txtDocumentNo9.Text = "";
            txtDocumentNo10.Text = "";
            txtDocumentNo11.Text = "";
            txtDocumentNo12.Text = "";
            txtDocumentNo13.Text = "";
            txtDocumentNo14.Text = "";
            txtDocumentNo15.Text = "";

            txtInvoiceNo1.Text = "";
            InvoiceDate1.Text = "";
            InvoiceValue1.Text = "";
            InvoiceUnit1.Text = "";

            txtInvoiceNo2.Text = "";
            InvoiceDate2.Text = "";
            InvoiceValue2.Text = "";
            InvoiceUnit2.Text = "";

            txtInvoiceNo3.Text = "";
            InvoiceDate3.Text = "";
            InvoiceValue3.Text = "";
            InvoiceUnit3.Text = "";

            txtInvoiceNo4.Text = "";
            InvoiceDate4.Text = "";
            InvoiceValue4.Text = "";
            InvoiceUnit4.Text = "";

            txtInvoiceNo5.Text = "";
            InvoiceDate5.Text = "";
            InvoiceValue5.Text = "";
            InvoiceUnit5.Text = "";

            txtInvoiceNo6.Text = "";
            InvoiceDate6.Text = "";
            InvoiceValue6.Text = "";
            InvoiceUnit6.Text = "";

            txtInvoiceNo7.Text = "";
            InvoiceDate7.Text = "";
            InvoiceValue7.Text = "";
            InvoiceUnit7.Text = "";

            txtInvoiceNo8.Text = "";
            InvoiceDate8.Text = "";
            InvoiceValue8.Text = "";
            InvoiceUnit8.Text = "";

            txtInvoiceNo9.Text = "";
            InvoiceDate9.Text = "";
            InvoiceValue9.Text = "";
            InvoiceUnit9.Text = "";

            txtInvoiceNo10.Text = "";
            InvoiceDate10.Text = "";
            InvoiceValue10.Text = "";
            InvoiceUnit10.Text = "";

            txtInvoiceNo11.Text = "";
            InvoiceDate11.Text = "";
            InvoiceValue11.Text = "";
            InvoiceUnit11.Text = "";

            txtInvoiceNo12.Text = "";
            InvoiceDate12.Text = "";
            InvoiceValue12.Text = "";
            InvoiceUnit12.Text = "";

            txtInvoiceNo13.Text = "";
            InvoiceDate13.Text = "";
            InvoiceValue13.Text = "";
            InvoiceUnit13.Text = "";

            txtInvoiceNo14.Text = "";
            InvoiceDate14.Text = "";
            InvoiceValue14.Text = "";
            InvoiceUnit14.Text = "";

            txtInvoiceNo15.Text = "";
            InvoiceDate15.Text = "";
            InvoiceValue15.Text = "";
            InvoiceUnit15.Text = "";

            txtEwayBillNo1.Text = "";
            txtEwayBillDate1.Text = "";
            txtValidUptoDate1.Text = "";


            txtEwayBillNo2.Text = "";
            txtEwayBillDate2.Text = "";
            txtValidUptoDate2.Text = "";


            txtEwayBillNo3.Text = "";
            txtEwayBillDate3.Text = "";
            txtValidUptoDate3.Text = "";


            txtEwayBillNo4.Text = "";
            txtEwayBillDate4.Text = "";
            txtValidUptoDate4.Text = "";


            txtEwayBillNo5.Text = "";
            txtEwayBillDate5.Text = "";
            txtValidUptoDate5.Text = "";

            txtEwayBillNo6.Text = "";
            txtEwayBillDate6.Text = "";
            txtValidUptoDate6.Text = "";

            txtEwayBillNo7.Text = "";
            txtEwayBillDate7.Text = "";
            txtValidUptoDate7.Text = "";

            txtEwayBillNo8.Text = "";
            txtEwayBillDate8.Text = "";
            txtValidUptoDate8.Text = "";

            txtEwayBillNo9.Text = "";
            txtEwayBillDate9.Text = "";
            txtValidUptoDate9.Text = "";

            txtEwayBillNo10.Text = "";
            txtEwayBillDate10.Text = "";
            txtValidUptoDate10.Text = "";

            txtEwayBillNo11.Text = "";
            txtEwayBillDate11.Text = "";
            txtValidUptoDate11.Text = "";

            txtEwayBillNo12.Text = "";
            txtEwayBillDate12.Text = "";
            txtValidUptoDate12.Text = "";

            txtEwayBillNo13.Text = "";
            txtEwayBillDate13.Text = "";
            txtValidUptoDate13.Text = "";

            txtEwayBillNo14.Text = "";
            txtEwayBillDate14.Text = "";
            txtValidUptoDate14.Text = "";

            txtEwayBillNo15.Text = "";
            txtEwayBillDate15.Text = "";
            txtValidUptoDate15.Text = "";
        }
        public string GetLastLrNo()
        {
            string LastLrNo = "";
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetLastLR");
                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    LastLrNo = dt.Rows[0]["BookLrNo"].ToString();
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return LastLrNo;
        }
        public int GetLRCount()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetLRCount");
                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public int GetLCCount()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetLCCount");
                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public int GetMnaualLRCount()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("GetManualLRCount");
                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                ParameterName.Add("@UserType"); ParameterValue.Add((cm.GetUserDetails().Rows[0]["UserType"]).ToString());
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public int CheckDuplicateLR()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("CheckDuplicate");
                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public int CheckDuplicateLR1()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("CheckDuplicate1");
                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                ParameterName.Add("@BookId"); ParameterValue.Add(HFid.Value);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }

        public int CheckDuplicateLC()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("CheckDuplicateLC");
                ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public int CheckDuplicateLC1()
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("CheckDuplicateLC1");
                ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                ParameterName.Add("@BookId"); ParameterValue.Add(HFid.Value);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return Convert.ToInt32(dt.Rows[0]["Count"]);
        }
        public void GetLRData(int BookId)
        {
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("Get");
                ParameterName.Add("@BookId"); ParameterValue.Add(BookId);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ddVType.SelectedValue = dt.Rows[0]["BookType"].ToString();
                    TxtLrNo.Text = dt.Rows[0]["BookLrNo"].ToString();
                    TxtLrNo.ReadOnly = true;
                    txtdate.Text = dt.Rows[0]["BookDate"].ToString();
                    ddSourceFrom.SelectedValue = dt.Rows[0]["SourceFromid"].ToString();
                    ddSourceTo.SelectedValue = dt.Rows[0]["SourceToid"].ToString();
                    txtDestination.Text = dt.Rows[0]["BookDestination"].ToString();
                    ddConsignor.SelectedValue = dt.Rows[0]["BookConsiginorid"].ToString();
                    ddlGSTPay.SelectedValue = dt.Rows[0]["GSTPay"].ToString();
                    ddConsignee.SelectedValue = dt.Rows[0]["BookConsigieeid"].ToString();
                    FillVehicleById(Convert.ToInt32(dt.Rows[0]["BookVchid"]));
                    txtinvno.Text = dt.Rows[0]["BookInvNo"].ToString();
                    txtinvdate.Text = dt.Rows[0]["BookInvDate"].ToString();
                    txtAmount.Text = dt.Rows[0]["BookAmount"].ToString();
                    txtEwayBillNo.Text = dt.Rows[0]["BookEwayBillNo"].ToString();
                    txtfromdate.Text = dt.Rows[0]["BookFromDate"].ToString();
                    txttodate.Text = dt.Rows[0]["BookValidDate"].ToString();
                    txtShipNo.Text = dt.Rows[0]["BookShipNo"].ToString();
                    txtDocNo.Text = dt.Rows[0]["BookDocNo"].ToString();
                    txtLoading.Text = dt.Rows[0]["BookLoading"].ToString();

                    txtRemarks.Text = dt.Rows[0]["BookRemarks"].ToString();
                    txtProduct.Text = dt.Rows[0]["ProdName"].ToString();
                    txtDriverName.Text = dt.Rows[0]["DriverName"].ToString();
                    ddlMode.SelectedValue = dt.Rows[0]["Mtype"].ToString();
                    ddlunit.SelectedValue = dt.Rows[0]["UnitName"].ToString();
                    ddlBillto.SelectedValue = dt.Rows[0]["BookBillTo"].ToString();
                    ddlSize.SelectedValue = dt.Rows[0]["BookCustSize"].ToString();
                    txtsealno.Text = dt.Rows[0]["SealNo"].ToString();
                    txtQty.Text = dt.Rows[0]["BookQty"].ToString();
                    hfbrid.Value = dt.Rows[0]["BookBranchid"].ToString();
                    txtDriverMobile.Text = dt.Rows[0]["BookDriverMobileNo"].ToString();
                    txtRemarks1.Text = dt.Rows[0]["BookRemark1"].ToString();
                    txtDriverAdhar.Text = dt.Rows[0]["DriverUID"].ToString();
                    txtLCNo.Text = dt.Rows[0]["LCNo"].ToString();
                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    spName = "Proc_DocumentDetails";

                    ParameterName.Add("@Action"); ParameterValue.Add("GetAll");
                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count == 1)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                    }
                    if (dt.Rows.Count == 2)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                    }
                    if (dt.Rows.Count == 3)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                    }
                    if (dt.Rows.Count == 4)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                    }
                    if (dt.Rows.Count == 5)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                    }
                    if (dt.Rows.Count == 6)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 7)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 8)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 9)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 10)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 11)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                        txtDocumentNo11.Text = dt.Rows[10]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 12)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                        txtDocumentNo11.Text = dt.Rows[10]["DocumentNo"].ToString();
                        txtDocumentNo12.Text = dt.Rows[11]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 13)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                        txtDocumentNo11.Text = dt.Rows[10]["DocumentNo"].ToString();
                        txtDocumentNo12.Text = dt.Rows[11]["DocumentNo"].ToString();
                        txtDocumentNo13.Text = dt.Rows[12]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 14)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                        txtDocumentNo11.Text = dt.Rows[10]["DocumentNo"].ToString();
                        txtDocumentNo12.Text = dt.Rows[11]["DocumentNo"].ToString();
                        txtDocumentNo13.Text = dt.Rows[12]["DocumentNo"].ToString();
                        txtDocumentNo14.Text = dt.Rows[13]["DocumentNo"].ToString();
                    }

                    if (dt.Rows.Count == 15)
                    {
                        txtDocumentNo1.Text = dt.Rows[0]["DocumentNo"].ToString();
                        txtDocumentNo2.Text = dt.Rows[1]["DocumentNo"].ToString();
                        txtDocumentNo3.Text = dt.Rows[2]["DocumentNo"].ToString();
                        txtDocumentNo4.Text = dt.Rows[3]["DocumentNo"].ToString();
                        txtDocumentNo5.Text = dt.Rows[4]["DocumentNo"].ToString();
                        txtDocumentNo6.Text = dt.Rows[5]["DocumentNo"].ToString();
                        txtDocumentNo7.Text = dt.Rows[6]["DocumentNo"].ToString();
                        txtDocumentNo8.Text = dt.Rows[7]["DocumentNo"].ToString();
                        txtDocumentNo9.Text = dt.Rows[8]["DocumentNo"].ToString();
                        txtDocumentNo10.Text = dt.Rows[9]["DocumentNo"].ToString();
                        txtDocumentNo11.Text = dt.Rows[10]["DocumentNo"].ToString();
                        txtDocumentNo12.Text = dt.Rows[11]["DocumentNo"].ToString();
                        txtDocumentNo13.Text = dt.Rows[12]["DocumentNo"].ToString();
                        txtDocumentNo14.Text = dt.Rows[13]["DocumentNo"].ToString();
                        txtDocumentNo15.Text = dt.Rows[14]["DocumentNo"].ToString();
                    }


                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    spName = "Proc_InvoiceDetails";

                    ParameterName.Add("@Action"); ParameterValue.Add("GetAll");
                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count == 1)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();
                    }
                    if (dt.Rows.Count == 2)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();
                    }
                    if (dt.Rows.Count == 3)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();
                    }

                    if (dt.Rows.Count == 4)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();
                    }
                    if (dt.Rows.Count == 5)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                     
                    }

                    if (dt.Rows.Count == 6)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                      
                    }

                    if (dt.Rows.Count == 7)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                      
                    }

                    if (dt.Rows.Count == 8)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                     
                    }

                    if (dt.Rows.Count == 9)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        
                    }

                    if (dt.Rows.Count == 10)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                 
                    }

                    if (dt.Rows.Count == 11)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                        txtInvoiceNo11.Text = dt.Rows[10]["InvoiceNo"].ToString();
                        InvoiceDate11.Text = dt.Rows[10]["InvoiceDate"].ToString();
                        InvoiceValue11.Text = dt.Rows[10]["InvoiceValue"].ToString();
                        InvoiceUnit11.Text = dt.Rows[10]["InvoiceUnit"].ToString();

                   
                    }

                    if (dt.Rows.Count == 12)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                        txtInvoiceNo11.Text = dt.Rows[10]["InvoiceNo"].ToString();
                        InvoiceDate11.Text = dt.Rows[10]["InvoiceDate"].ToString();
                        InvoiceValue11.Text = dt.Rows[10]["InvoiceValue"].ToString();
                        InvoiceUnit11.Text = dt.Rows[10]["InvoiceUnit"].ToString();

                        txtInvoiceNo12.Text = dt.Rows[11]["InvoiceNo"].ToString();
                        InvoiceDate12.Text = dt.Rows[11]["InvoiceDate"].ToString();
                        InvoiceValue12.Text = dt.Rows[11]["InvoiceValue"].ToString();
                        InvoiceUnit12.Text = dt.Rows[11]["InvoiceUnit"].ToString();

                        
                    }
                    if (dt.Rows.Count == 13)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                        txtInvoiceNo11.Text = dt.Rows[10]["InvoiceNo"].ToString();
                        InvoiceDate11.Text = dt.Rows[10]["InvoiceDate"].ToString();
                        InvoiceValue11.Text = dt.Rows[10]["InvoiceValue"].ToString();
                        InvoiceUnit11.Text = dt.Rows[10]["InvoiceUnit"].ToString();

                        txtInvoiceNo12.Text = dt.Rows[11]["InvoiceNo"].ToString();
                        InvoiceDate12.Text = dt.Rows[11]["InvoiceDate"].ToString();
                        InvoiceValue12.Text = dt.Rows[11]["InvoiceValue"].ToString();
                        InvoiceUnit12.Text = dt.Rows[11]["InvoiceUnit"].ToString();

                        txtInvoiceNo13.Text = dt.Rows[12]["InvoiceNo"].ToString();
                        InvoiceDate13.Text = dt.Rows[12]["InvoiceDate"].ToString();
                        InvoiceValue13.Text = dt.Rows[12]["InvoiceValue"].ToString();
                        InvoiceUnit13.Text = dt.Rows[12]["InvoiceUnit"].ToString();

                    
                    }

                    if (dt.Rows.Count == 14)
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                        txtInvoiceNo11.Text = dt.Rows[10]["InvoiceNo"].ToString();
                        InvoiceDate11.Text = dt.Rows[10]["InvoiceDate"].ToString();
                        InvoiceValue11.Text = dt.Rows[10]["InvoiceValue"].ToString();
                        InvoiceUnit11.Text = dt.Rows[10]["InvoiceUnit"].ToString();

                        txtInvoiceNo12.Text = dt.Rows[11]["InvoiceNo"].ToString();
                        InvoiceDate12.Text = dt.Rows[11]["InvoiceDate"].ToString();
                        InvoiceValue12.Text = dt.Rows[11]["InvoiceValue"].ToString();
                        InvoiceUnit12.Text = dt.Rows[11]["InvoiceUnit"].ToString();

                        txtInvoiceNo13.Text = dt.Rows[12]["InvoiceNo"].ToString();
                        InvoiceDate13.Text = dt.Rows[12]["InvoiceDate"].ToString();
                        InvoiceValue13.Text = dt.Rows[12]["InvoiceValue"].ToString();
                        InvoiceUnit13.Text = dt.Rows[12]["InvoiceUnit"].ToString();

                        txtInvoiceNo14.Text = dt.Rows[13]["InvoiceNo"].ToString();
                        InvoiceDate14.Text = dt.Rows[13]["InvoiceDate"].ToString();
                        InvoiceValue14.Text = dt.Rows[13]["InvoiceValue"].ToString();
                        InvoiceUnit14.Text = dt.Rows[13]["InvoiceUnit"].ToString();

                    }

                    if (dt.Rows.Count == 15)    
                    {
                        txtInvoiceNo1.Text = dt.Rows[0]["InvoiceNo"].ToString();
                        InvoiceDate1.Text = dt.Rows[0]["InvoiceDate"].ToString();
                        InvoiceValue1.Text = dt.Rows[0]["InvoiceValue"].ToString();
                        InvoiceUnit1.Text = dt.Rows[0]["InvoiceUnit"].ToString();

                        txtInvoiceNo2.Text = dt.Rows[1]["InvoiceNo"].ToString();
                        InvoiceDate2.Text = dt.Rows[1]["InvoiceDate"].ToString();
                        InvoiceValue2.Text = dt.Rows[1]["InvoiceValue"].ToString();
                        InvoiceUnit2.Text = dt.Rows[1]["InvoiceUnit"].ToString();

                        txtInvoiceNo3.Text = dt.Rows[2]["InvoiceNo"].ToString();
                        InvoiceDate3.Text = dt.Rows[2]["InvoiceDate"].ToString();
                        InvoiceValue3.Text = dt.Rows[2]["InvoiceValue"].ToString();
                        InvoiceUnit3.Text = dt.Rows[2]["InvoiceUnit"].ToString();

                        txtInvoiceNo4.Text = dt.Rows[3]["InvoiceNo"].ToString();
                        InvoiceDate4.Text = dt.Rows[3]["InvoiceDate"].ToString();
                        InvoiceValue4.Text = dt.Rows[3]["InvoiceValue"].ToString();
                        InvoiceUnit4.Text = dt.Rows[3]["InvoiceUnit"].ToString();

                        txtInvoiceNo5.Text = dt.Rows[4]["InvoiceNo"].ToString();
                        InvoiceDate5.Text = dt.Rows[4]["InvoiceDate"].ToString();
                        InvoiceValue5.Text = dt.Rows[4]["InvoiceValue"].ToString();
                        InvoiceUnit5.Text = dt.Rows[4]["InvoiceUnit"].ToString();

                        txtInvoiceNo6.Text = dt.Rows[5]["InvoiceNo"].ToString();
                        InvoiceDate6.Text = dt.Rows[5]["InvoiceDate"].ToString();
                        InvoiceValue6.Text = dt.Rows[5]["InvoiceValue"].ToString();
                        InvoiceUnit6.Text = dt.Rows[5]["InvoiceUnit"].ToString();

                        txtInvoiceNo7.Text = dt.Rows[6]["InvoiceNo"].ToString();
                        InvoiceDate7.Text = dt.Rows[6]["InvoiceDate"].ToString();
                        InvoiceValue7.Text = dt.Rows[6]["InvoiceValue"].ToString();
                        InvoiceUnit7.Text = dt.Rows[6]["InvoiceUnit"].ToString();

                        txtInvoiceNo8.Text = dt.Rows[7]["InvoiceNo"].ToString();
                        InvoiceDate8.Text = dt.Rows[7]["InvoiceDate"].ToString();
                        InvoiceValue8.Text = dt.Rows[7]["InvoiceValue"].ToString();
                        InvoiceUnit8.Text = dt.Rows[7]["InvoiceUnit"].ToString();

                        txtInvoiceNo9.Text = dt.Rows[8]["InvoiceNo"].ToString();
                        InvoiceDate9.Text = dt.Rows[8]["InvoiceDate"].ToString();
                        InvoiceValue9.Text = dt.Rows[8]["InvoiceValue"].ToString();
                        InvoiceUnit9.Text = dt.Rows[8]["InvoiceUnit"].ToString();

                        txtInvoiceNo10.Text = dt.Rows[9]["InvoiceNo"].ToString();
                        InvoiceDate10.Text = dt.Rows[9]["InvoiceDate"].ToString();
                        InvoiceValue10.Text = dt.Rows[9]["InvoiceValue"].ToString();
                        InvoiceUnit10.Text = dt.Rows[9]["InvoiceUnit"].ToString();

                        txtInvoiceNo11.Text = dt.Rows[10]["InvoiceNo"].ToString();
                        InvoiceDate11.Text = dt.Rows[10]["InvoiceDate"].ToString();
                        InvoiceValue11.Text = dt.Rows[10]["InvoiceValue"].ToString();
                        InvoiceUnit11.Text = dt.Rows[10]["InvoiceUnit"].ToString();

                        txtInvoiceNo12.Text = dt.Rows[11]["InvoiceNo"].ToString();
                        InvoiceDate12.Text = dt.Rows[11]["InvoiceDate"].ToString();
                        InvoiceValue12.Text = dt.Rows[11]["InvoiceValue"].ToString();
                        InvoiceUnit12.Text = dt.Rows[11]["InvoiceUnit"].ToString();

                        txtInvoiceNo13.Text = dt.Rows[12]["InvoiceNo"].ToString();
                        InvoiceDate13.Text = dt.Rows[12]["InvoiceDate"].ToString();
                        InvoiceValue13.Text = dt.Rows[12]["InvoiceValue"].ToString();
                        InvoiceUnit13.Text = dt.Rows[12]["InvoiceUnit"].ToString();

                        txtInvoiceNo14.Text = dt.Rows[13]["InvoiceNo"].ToString();
                        InvoiceDate14.Text = dt.Rows[13]["InvoiceDate"].ToString();
                        InvoiceValue14.Text = dt.Rows[13]["InvoiceValue"].ToString();
                        InvoiceUnit14.Text = dt.Rows[13]["InvoiceUnit"].ToString();

                        txtInvoiceNo15.Text = dt.Rows[14]["InvoiceNo"].ToString();
                        InvoiceDate15.Text = dt.Rows[14]["InvoiceDate"].ToString();
                        InvoiceValue15.Text = dt.Rows[14]["InvoiceValue"].ToString();
                        InvoiceUnit15.Text = dt.Rows[14]["InvoiceUnit"].ToString();
                    }

                    ParameterName = new ArrayList();
                    ParameterValue = new ArrayList();
                    spName = "Proc_EwayBillDetails";

                    ParameterName.Add("@Action"); ParameterValue.Add("GetAll");
                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count == 1)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 2)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 3)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 4)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 5)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 6)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();
                    }

                    if (dt.Rows.Count == 7)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();
                    }
                    if (dt.Rows.Count == 8)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                       
                    }

                    if (dt.Rows.Count == 9)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                       
                    }

                    if (dt.Rows.Count == 10)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        
                    }

                    if (dt.Rows.Count == 11)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        txtEwayBillNo11.Text = dt.Rows[10]["EwayBillNo"].ToString();
                        txtEwayBillDate11.Text = dt.Rows[10]["BillDate"].ToString();
                        txtValidUptoDate11.Text = dt.Rows[10]["BillUptoDate"].ToString();

                       
                    }

                    if (dt.Rows.Count == 12)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        txtEwayBillNo11.Text = dt.Rows[10]["EwayBillNo"].ToString();
                        txtEwayBillDate11.Text = dt.Rows[10]["BillDate"].ToString();
                        txtValidUptoDate11.Text = dt.Rows[10]["BillUptoDate"].ToString();

                        txtEwayBillNo12.Text = dt.Rows[11]["EwayBillNo"].ToString();
                        txtEwayBillDate12.Text = dt.Rows[11]["BillDate"].ToString();
                        txtValidUptoDate12.Text = dt.Rows[11]["BillUptoDate"].ToString();

                       
                    }

                    if (dt.Rows.Count == 13)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        txtEwayBillNo11.Text = dt.Rows[10]["EwayBillNo"].ToString();
                        txtEwayBillDate11.Text = dt.Rows[10]["BillDate"].ToString();
                        txtValidUptoDate11.Text = dt.Rows[10]["BillUptoDate"].ToString();

                        txtEwayBillNo12.Text = dt.Rows[11]["EwayBillNo"].ToString();
                        txtEwayBillDate12.Text = dt.Rows[11]["BillDate"].ToString();
                        txtValidUptoDate12.Text = dt.Rows[11]["BillUptoDate"].ToString();

                        txtEwayBillNo13.Text = dt.Rows[12]["EwayBillNo"].ToString();
                        txtEwayBillDate13.Text = dt.Rows[12]["BillDate"].ToString();
                        txtValidUptoDate13.Text = dt.Rows[12]["BillUptoDate"].ToString();

                       
                    }

                    if (dt.Rows.Count == 14)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        txtEwayBillNo11.Text = dt.Rows[10]["EwayBillNo"].ToString();
                        txtEwayBillDate11.Text = dt.Rows[10]["BillDate"].ToString();
                        txtValidUptoDate11.Text = dt.Rows[10]["BillUptoDate"].ToString();

                        txtEwayBillNo12.Text = dt.Rows[11]["EwayBillNo"].ToString();
                        txtEwayBillDate12.Text = dt.Rows[11]["BillDate"].ToString();
                        txtValidUptoDate12.Text = dt.Rows[11]["BillUptoDate"].ToString();

                        txtEwayBillNo13.Text = dt.Rows[12]["EwayBillNo"].ToString();
                        txtEwayBillDate13.Text = dt.Rows[12]["BillDate"].ToString();
                        txtValidUptoDate13.Text = dt.Rows[12]["BillUptoDate"].ToString();

                        txtEwayBillNo14.Text = dt.Rows[13]["EwayBillNo"].ToString();
                        txtEwayBillDate14.Text = dt.Rows[13]["BillDate"].ToString();
                        txtValidUptoDate14.Text = dt.Rows[13]["BillUptoDate"].ToString();

                     
                    }

                    if (dt.Rows.Count == 15)
                    {
                        txtEwayBillNo1.Text = dt.Rows[0]["EwayBillNo"].ToString();
                        txtEwayBillDate1.Text = dt.Rows[0]["BillDate"].ToString();
                        txtValidUptoDate1.Text = dt.Rows[0]["BillUptoDate"].ToString();

                        txtEwayBillNo2.Text = dt.Rows[1]["EwayBillNo"].ToString();
                        txtEwayBillDate2.Text = dt.Rows[1]["BillDate"].ToString();
                        txtValidUptoDate2.Text = dt.Rows[1]["BillUptoDate"].ToString();

                        txtEwayBillNo3.Text = dt.Rows[2]["EwayBillNo"].ToString();
                        txtEwayBillDate3.Text = dt.Rows[2]["BillDate"].ToString();
                        txtValidUptoDate3.Text = dt.Rows[2]["BillUptoDate"].ToString();

                        txtEwayBillNo4.Text = dt.Rows[3]["EwayBillNo"].ToString();
                        txtEwayBillDate4.Text = dt.Rows[3]["BillDate"].ToString();
                        txtValidUptoDate4.Text = dt.Rows[3]["BillUptoDate"].ToString();

                        txtEwayBillNo5.Text = dt.Rows[4]["EwayBillNo"].ToString();
                        txtEwayBillDate5.Text = dt.Rows[4]["BillDate"].ToString();
                        txtValidUptoDate5.Text = dt.Rows[4]["BillUptoDate"].ToString();

                        txtEwayBillNo6.Text = dt.Rows[5]["EwayBillNo"].ToString();
                        txtEwayBillDate6.Text = dt.Rows[5]["BillDate"].ToString();
                        txtValidUptoDate6.Text = dt.Rows[5]["BillUptoDate"].ToString();

                        txtEwayBillNo7.Text = dt.Rows[6]["EwayBillNo"].ToString();
                        txtEwayBillDate7.Text = dt.Rows[6]["BillDate"].ToString();
                        txtValidUptoDate7.Text = dt.Rows[6]["BillUptoDate"].ToString();

                        txtEwayBillNo8.Text = dt.Rows[7]["EwayBillNo"].ToString();
                        txtEwayBillDate8.Text = dt.Rows[7]["BillDate"].ToString();
                        txtValidUptoDate8.Text = dt.Rows[7]["BillUptoDate"].ToString();

                        txtEwayBillNo9.Text = dt.Rows[8]["EwayBillNo"].ToString();
                        txtEwayBillDate9.Text = dt.Rows[8]["BillDate"].ToString();
                        txtValidUptoDate9.Text = dt.Rows[8]["BillUptoDate"].ToString();

                        txtEwayBillNo10.Text = dt.Rows[9]["EwayBillNo"].ToString();
                        txtEwayBillDate10.Text = dt.Rows[9]["BillDate"].ToString();
                        txtValidUptoDate10.Text = dt.Rows[9]["BillUptoDate"].ToString();

                        txtEwayBillNo11.Text = dt.Rows[10]["EwayBillNo"].ToString();
                        txtEwayBillDate11.Text = dt.Rows[10]["BillDate"].ToString();
                        txtValidUptoDate11.Text = dt.Rows[10]["BillUptoDate"].ToString();

                        txtEwayBillNo12.Text = dt.Rows[11]["EwayBillNo"].ToString();
                        txtEwayBillDate12.Text = dt.Rows[11]["BillDate"].ToString();
                        txtValidUptoDate12.Text = dt.Rows[11]["BillUptoDate"].ToString();

                        txtEwayBillNo13.Text = dt.Rows[12]["EwayBillNo"].ToString();
                        txtEwayBillDate13.Text = dt.Rows[12]["BillDate"].ToString();
                        txtValidUptoDate13.Text = dt.Rows[12]["BillUptoDate"].ToString();

                        txtEwayBillNo14.Text = dt.Rows[13]["EwayBillNo"].ToString();
                        txtEwayBillDate14.Text = dt.Rows[13]["BillDate"].ToString();
                        txtValidUptoDate14.Text = dt.Rows[13]["BillUptoDate"].ToString();

                        txtEwayBillNo15.Text = dt.Rows[14]["EwayBillNo"].ToString();
                        txtEwayBillDate15.Text = dt.Rows[14]["BillDate"].ToString();
                        txtValidUptoDate15.Text = dt.Rows[14]["BillUptoDate"].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        public string MaxBookId()
        {
            string BookId = "";
            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_BookLR";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("MaxBookId");
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    BookId = dt.Rows[0]["BookId"].ToString();
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
            return BookId;
        }

        // Multi-user safe: is user ne jo LR abhi save kiya, sirf usi ka BookId lao
        // (MaxBookId() poore table ka latest deta hai - kisi aur user ka LR bhi ho sakta hai)
        public string GetBookIdByLrNo(string LrNo)
        {
            return GetSingleBookId("SELECT MAX(BookId) AS BookId FROM tblBookLR WHERE BookLrNo = @LrNo", "@LrNo", LrNo);
        }

        public string GetBookIdByGroupId(long GroupId)
        {
            return GetSingleBookId("SELECT MAX(BookId) AS BookId FROM tblBookLR WHERE GroupId = @GroupId", "@GroupId", GroupId.ToString());
        }

        private string GetSingleBookId(string query, string paramName, string paramValue)
        {
            string BookId = "";
            try
            {
                ArrayList pn = new ArrayList(); ArrayList pv = new ArrayList();
                pn.Add(paramName); pv.Add(paramValue);
                DataTable t = con.GetDataTableFromTable(query, pn, pv);
                if (t.Rows.Count > 0 && t.Rows[0]["BookId"] != DBNull.Value)
                {
                    BookId = t.Rows[0]["BookId"].ToString();
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:GetSingleBookId | SP: | User:" + Convert.ToString(Session["UserName"]) + " | Err:" + ex.Message.ToString());
            }
            return BookId;
        }

        // Multi-LR batch ke liye unique GroupId (pehle 5-digit OTP use hota tha jo users ke beech repeat ho sakta tha)
        private string NewGroupId()
        {
            int userId = Convert.ToInt32(Session["UserId"]) % 1000;
            return DateTime.Now.ToString("yyMMddHHmmssfff") + userId.ToString("000");
        }
        #endregion

        #region[Events]
        public int CheckInvoiceValue()
        {
            int check = 0;
            if (InvoiceValue1.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue1.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue2.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue2.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue3.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue3.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue4.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue4.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue5.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue5.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue6.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue6.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            if (InvoiceValue7.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue7.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue8.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue8.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue9.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue9.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue10.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue10.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue11.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue11.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue12.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue12.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue13.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue13.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue14.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue14.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }

            if (InvoiceValue15.Text != "")
            {
                if (!(Regex.IsMatch(InvoiceValue15.Text.Trim(), @"^\d+(,\d+)*(\.\d+)?$")))
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Invoice Value not in correct format');", true);
                    check = 1;
                }
            }
            return check;
        }
        protected async void btnAdd_Click(object sender, EventArgs e)
        {
            int Count = 0;
            int Count1 = 0;
            string sSql = "";
            CultureInfo cultuers = new CultureInfo("en-US");
            Int32 IntResult = 0;
            string spName = "";
            try
            {
                if (Session["OTP"] == null)
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('OTP expired. Please generate OTP again and save.');", true);
                    return;
                }
                if (txtCode.Text == Session["OTP"].ToString())
                {
                    if (CheckInvoiceValue() == 1)
                    {
                        return;
                    }
                    if (ddlMode.SelectedIndex == 2 && HFid.Value == "")
                    {
                        if (Session["GroupSessionId"] == null)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to add atleast one LR');", true);
                        }
                        else
                        {
                            ParameterName = new ArrayList();
                            ParameterValue = new ArrayList();
                            spName = "Proc_BookLR";
                            ParameterName.Add("@Action"); ParameterValue.Add("MoveFromTMPtoMain");
                            ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                            IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            if (IntResult > 0)
                            {
                                string GroupBookId = GetBookIdByGroupId(Convert.ToInt64(Session["GroupSessionId"]));
                                Session["GroupSessionId"] = null;
                                Page.ClientScript.RegisterStartupScript(this.GetType(), Guid.NewGuid().ToString(), "alert('All Data SuccessFully inserted!');window.location='ReportViewer.aspx?BookId=" + GroupBookId + "';", true);
                                Session["PrintBookId"] = GroupBookId;
                                Clear();
                            }
                            else
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('No pending LR found for this batch. Please check LR list before re-entering.');", true);
                            }
                        }
                    }
                    else
                    {
                        if (ddlMode.SelectedIndex == 0)
                        {
                            Count = GetLRCount();
                        }
                        else
                        {
                            Count = GetMnaualLRCount();
                        }

                        
                        Count1 = GetLCCount();

                        if (HFid.Value == "")
                        {
                            if (CheckDuplicateLR() > 0)
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LR No!');", true);
                            }
                            else if (Count == 0)
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for Consignment No!');TxtLrNo.focus;", true);
                            }
                            //else if (CheckDuplicateLC() > 0)
                            //{
                            //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LC No!');", true);
                            //}
                            else if (Count1 == 0 && txtLCNo.Text != "")
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('please contact to admin for lc no!');txtlcno.focus;", true);
                            }
                            //else if (ddlMode.SelectedIndex == 0 && TxtLrNo.Text.Length != 7)
                            //{
                            //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Consigment No Must be in 7 Digits!');TxtLrNo.focus;", true);
                            //}
                            else
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                spName = "Proc_BookLR";
                                ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                ParameterName.Add("@BookType"); ParameterValue.Add(ddVType.SelectedValue);
                                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@BookDate"); ParameterValue.Add(txtdate.Text);
                                ParameterName.Add("@SourceFromid"); ParameterValue.Add(ddSourceFrom.SelectedValue);
                                ParameterName.Add("@SourceToid"); ParameterValue.Add(ddSourceTo.SelectedValue);
                                ParameterName.Add("@BookDestination"); ParameterValue.Add(txtDestination.Text);
                                ParameterName.Add("@BookConsiginorid"); ParameterValue.Add(ddConsignor.SelectedValue);
                                ParameterName.Add("@BookConsigieeid"); ParameterValue.Add(ddConsignee.SelectedValue);
                                ParameterName.Add("@BookPartyid"); ParameterValue.Add(0);
                                ParameterName.Add("@BookVchid"); ParameterValue.Add(ddlVchNo.SelectedValue);
                                ParameterName.Add("@BookInvNo"); ParameterValue.Add(txtinvno.Text);
                                ParameterName.Add("@BookInvDate"); ParameterValue.Add(txtinvdate.Text);
                                ParameterName.Add("@BookAmount"); ParameterValue.Add(txtAmount.Text);
                                ParameterName.Add("@BookEwayBillNo"); ParameterValue.Add(txtEwayBillNo.Text);
                                ParameterName.Add("@BookFromDate"); ParameterValue.Add(txtfromdate.Text);
                                ParameterName.Add("@BookValidDate"); ParameterValue.Add(txttodate.Text);
                                ParameterName.Add("@BookShipNo"); ParameterValue.Add(txtShipNo.Text);
                                ParameterName.Add("@BookDocNo"); ParameterValue.Add(txtDocNo.Text);
                                ParameterName.Add("@BookLoading"); ParameterValue.Add(txtLoading.Text);
                                ParameterName.Add("@GSTPay"); ParameterValue.Add(ddlGSTPay.SelectedValue);
                                ParameterName.Add("@BookRemarks"); ParameterValue.Add(txtRemarks.Text);
                                ParameterName.Add("@ProdName"); ParameterValue.Add(txtProduct.Text.Trim()); SaveProduct(txtProduct.Text.Trim());
                                ParameterName.Add("@UnitName"); ParameterValue.Add(ddlunit.SelectedValue);
                                ParameterName.Add("@BookQty"); ParameterValue.Add(txtQty.Text);
                                ParameterName.Add("@ActQty"); ParameterValue.Add(0);
                                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                                ParameterName.Add("@BookDriverMobileNo"); ParameterValue.Add(txtDriverMobile.Text);
                                ParameterName.Add("@BookCustSize"); ParameterValue.Add(ddlSize.SelectedValue);
                                ParameterName.Add("@BookRemark1"); ParameterValue.Add(txtRemarks1.Text);
                                ParameterName.Add("@DriverName"); ParameterValue.Add(txtDriverName.Text);
                                ParameterName.Add("@SealNo"); ParameterValue.Add(txtsealno.Text);
                                ParameterName.Add("@BookBillTo"); ParameterValue.Add(ddlBillto.SelectedValue);
                                ParameterName.Add("@Mtype"); ParameterValue.Add(ddlMode.SelectedValue);
                                ParameterName.Add("@UserId"); ParameterValue.Add(Convert.ToInt32(Session["UserId"]));
                                ParameterName.Add("@DriverUID"); ParameterValue.Add(txtDriverAdhar.Text);
                                ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                if (IntResult > 0)
                                {
                                    string MaxBookId1 = GetBookIdByLrNo(TxtLrNo.Text);
                                    spName = "Proc_DocumentDetails";

                                    if (txtDocumentNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    spName = "Proc_InvoiceDetails";
                                    if (txtInvoiceNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo1.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate1.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue1.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo2.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate2.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue2.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo3.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate3.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue3.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo4.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate4.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue4.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo5.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate5.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue5.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo6.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate6.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue6.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo7.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate7.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue7.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo8.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate8.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue8.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo9.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate9.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue9.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo10.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate10.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue10.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo11.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate11.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue11.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo12.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate12.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue12.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo13.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate13.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue13.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo14.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate14.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue14.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo15.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate15.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue15.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    spName = "Proc_EwayBillDetails";
                                    if (txtEwayBillNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo1.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate1.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo2.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate2.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo3.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate3.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo4.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate4.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo5.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate5.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo6.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate6.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo7.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate7.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo8.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate8.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo9.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate9.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo10.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate10.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo11.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate11.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo12.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate12.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo13.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate13.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo14.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate14.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo15.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate15.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    //LC Entry
                                    ParameterName = new ArrayList();
                                    ParameterValue = new ArrayList();
                                    spName = "Proc_BookLR";
                                    ParameterName.Add("@Action"); ParameterValue.Add("LCbyLR");
                                    ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                                    ParameterName.Add("@BookDate"); ParameterValue.Add(txtdate.Text);
                                    ParameterName.Add("@SourceFromid"); ParameterValue.Add(ddSourceFrom.SelectedValue);
                                    ParameterName.Add("@SourceToid"); ParameterValue.Add(ddSourceTo.SelectedValue);
                                    ParameterName.Add("@BookDestination"); ParameterValue.Add(txtDestination.Text);
                                    ParameterName.Add("@BookVchid"); ParameterValue.Add(ddlVchNo.SelectedValue);
                                    ParameterName.Add("@UnitName"); ParameterValue.Add(ddlunit.SelectedValue);
                                    ParameterName.Add("@BookQty"); ParameterValue.Add(txtQty.Text);
                                    ParameterName.Add("@BookCustSize"); ParameterValue.Add(ddlSize.SelectedValue);
                                    ParameterName.Add("@UserId"); ParameterValue.Add(Convert.ToInt32(Session["UserId"]));
                                    ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                                    ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                                    ParameterName.Add("@DriverName"); ParameterValue.Add(txtDriverName.Text);
                                    ParameterName.Add("@DriverUID"); ParameterValue.Add(txtDriverAdhar.Text);
                                    ParameterName.Add("@BookDriverMobileNo"); ParameterValue.Add(txtDriverMobile.Text);
                                    IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);

                                    IntResult = cm.SaveActivityLog("LR Entry", "Inserted", TxtLrNo.Text);



                                    Page.ClientScript.RegisterStartupScript(this.GetType(), Guid.NewGuid().ToString(), "alert('Data SuccessFully inserted!');window.location='ReportViewer.aspx?BookId=" + MaxBookId1 + "';", true);
                                    if (txtEwayBillNo.Text.Trim() != "")
                                    {
                                        Session["PrintBookId"] = MaxBookId1;
                                    }


                                    GTropyAPIService service = new GTropyAPIService();
                                    await service.SendToGtropyAPI(ddlVchNo.SelectedItem.Text, TxtLrNo.Text, Convert.ToDateTime(txtdate.Text), ddSourceFrom.SelectedItem.Text, ddSourceTo.SelectedItem.Text);



                                    Clear();


                                }
                            }
                        }
                        else
                        {
                            if (CheckDuplicateLR1() > 0)
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LR No!');", true);
                            }
                            else if (Count == 0)
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for Consignment No!');TxtLrNo.focus;", true);
                            }
                            //else if (CheckDuplicateLC1() > 0)
                            //{
                            //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LC No!');", true);
                            //}
                            else if (Count1 == 0 && txtLCNo.Text != "")
                            {
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for LC No!');txtLCNo.focus;", true);
                            }
                            else
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                spName = "Proc_BookLR";
                                ParameterName.Add("@Action"); ParameterValue.Add("Update");
                                ParameterName.Add("@BookId"); ParameterValue.Add(HFid.Value);
                                ParameterName.Add("@BookType"); ParameterValue.Add(ddVType.SelectedValue);
                                ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@BookDate"); ParameterValue.Add(txtdate.Text);
                                ParameterName.Add("@SourceFromid"); ParameterValue.Add(ddSourceFrom.SelectedValue);
                                ParameterName.Add("@SourceToid"); ParameterValue.Add(ddSourceTo.SelectedValue);
                                ParameterName.Add("@BookDestination"); ParameterValue.Add(txtDestination.Text);
                                ParameterName.Add("@BookConsiginorid"); ParameterValue.Add(ddConsignor.SelectedValue);
                                ParameterName.Add("@BookConsigieeid"); ParameterValue.Add(ddConsignee.SelectedValue);
                                ParameterName.Add("@GSTPay"); ParameterValue.Add(ddlGSTPay.SelectedValue);
                                ParameterName.Add("@BookPartyid"); ParameterValue.Add(0);
                                ParameterName.Add("@BookVchid"); ParameterValue.Add(ddlVchNo.SelectedValue);
                                ParameterName.Add("@BookInvNo"); ParameterValue.Add(txtinvno.Text);
                                ParameterName.Add("@BookInvDate"); ParameterValue.Add(txtinvdate.Text);
                                ParameterName.Add("@BookAmount"); ParameterValue.Add(txtAmount.Text);
                                ParameterName.Add("@BookEwayBillNo"); ParameterValue.Add(txtEwayBillNo.Text);
                                ParameterName.Add("@BookFromDate"); ParameterValue.Add(txtfromdate.Text);
                                ParameterName.Add("@BookValidDate"); ParameterValue.Add(txttodate.Text);
                                ParameterName.Add("@BookShipNo"); ParameterValue.Add(txtShipNo.Text);
                                ParameterName.Add("@BookDocNo"); ParameterValue.Add(txtDocNo.Text);
                                ParameterName.Add("@BookLoading"); ParameterValue.Add(txtLoading.Text);
                                ParameterName.Add("@BookRemarks"); ParameterValue.Add(txtRemarks.Text);
                                ParameterName.Add("@ProdName"); ParameterValue.Add(txtProduct.Text.Trim()); SaveProduct(txtProduct.Text.Trim());
                                ParameterName.Add("@UnitName"); ParameterValue.Add(ddlunit.SelectedValue);
                                ParameterName.Add("@BookQty"); ParameterValue.Add(txtQty.Text);
                                ParameterName.Add("@ActQty"); ParameterValue.Add(0);
                                ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                                ParameterName.Add("@BookDriverMobileNo"); ParameterValue.Add(txtDriverMobile.Text);
                                ParameterName.Add("@BookCustSize"); ParameterValue.Add(ddlSize.SelectedValue);
                                ParameterName.Add("@BookRemark1"); ParameterValue.Add(txtRemarks1.Text);
                                ParameterName.Add("@DriverName"); ParameterValue.Add(txtDriverName.Text);
                                ParameterName.Add("@SealNo"); ParameterValue.Add(txtsealno.Text);
                                ParameterName.Add("@BookBillTo"); ParameterValue.Add(ddlBillto.SelectedValue);
                                ParameterName.Add("@Mtype"); ParameterValue.Add(ddlMode.SelectedValue);
                                ParameterName.Add("@UserId"); ParameterValue.Add(Convert.ToInt32(Session["UserId"]));
                                ParameterName.Add("@DriverUID"); ParameterValue.Add(txtDriverAdhar.Text);
                                ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                if (IntResult > 0)
                                {
                                    ParameterName = new ArrayList();
                                    ParameterValue = new ArrayList();
                                    spName = "Proc_DocumentDetails";

                                    ParameterName.Add("@Action"); ParameterValue.Add("delete");
                                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                    IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);

                                    ParameterName = new ArrayList();
                                    ParameterValue = new ArrayList();
                                    spName = "Proc_InvoiceDetails";

                                    ParameterName.Add("@Action"); ParameterValue.Add("delete");
                                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                    IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);

                                    ParameterName = new ArrayList();
                                    ParameterValue = new ArrayList();
                                    spName = "Proc_EwayBillDetails";

                                    ParameterName.Add("@Action"); ParameterValue.Add("delete");
                                    ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                    IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);


                                    spName = "Proc_DocumentDetails";
                                    if (txtDocumentNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtDocumentNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    
                                    if (txtDocumentNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtDocumentNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    spName = "Proc_InvoiceDetails";
                                    
                                    if (txtInvoiceNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo1.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate1.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue1.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo2.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate2.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue2.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo3.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate3.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue3.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo4.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate4.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue4.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo5.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate5.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue5.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtInvoiceNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo6.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate6.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue6.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo7.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate7.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue7.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo8.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate8.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue8.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo9.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate9.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue9.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo10.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate10.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue10.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo11.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate11.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue11.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo12.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate12.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue12.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo13.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate13.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue13.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo14.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate14.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue14.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtInvoiceNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo15.Text);
                                        ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate15.Text);
                                        ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue15.Text);
                                        ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }


                                    spName = "Proc_EwayBillDetails";
                                    if (txtEwayBillNo1.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo1.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate1.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate1.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo2.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo2.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate2.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate2.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo3.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo3.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate3.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate3.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo4.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo4.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate4.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate4.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo5.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo5.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate5.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate5.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }
                                    if (txtEwayBillNo6.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo6.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate6.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate6.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo7.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo7.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate7.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate7.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo8.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo8.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate8.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate8.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo9.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo9.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate9.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate9.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo10.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo10.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate10.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate10.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo11.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo11.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate11.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate11.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo12.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo12.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate12.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate12.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo13.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo13.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate13.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate13.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo14.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo14.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate14.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate14.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    if (txtEwayBillNo15.Text != "")
                                    {
                                        ParameterName = new ArrayList();
                                        ParameterValue = new ArrayList();
                                        ParameterName.Add("@Action"); ParameterValue.Add("Insert");
                                        ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                        ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo15.Text);
                                        ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate15.Text);
                                        ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate15.Text);
                                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                                    }

                                    IntResult = cm.SaveActivityLog("LR Entry", "Updated", TxtLrNo.Text);
                                    Page.ClientScript.RegisterStartupScript(this.GetType(), Guid.NewGuid().ToString(), "alert('Data SuccessFully Updated!');window.location='ReportViewer.aspx?BookId=" + HFid.Value + "';", true);
                                    if (txtEwayBillNo.Text.Trim() != "")
                                    {
                                        Session["PrintBookId"] = HFid.Value;
                                    }
                                    Clear();
                                }
                            }
                        }
                    }
                }
                else
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('In Correct OTP!');", true);
                    txtCode.Text = "";
                    divsave.Visible = true;
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        protected void btnSaveLocal_Click(object sender, EventArgs e)
        {
            int Count = 0;
            int Count1 = 0;
            CultureInfo cultuers = new CultureInfo("en-US");
            Int32 IntResult = 0;
            string OTPNumb = cm.FiveDigits();
            string spName = "";
            try
            {
                if (CheckInvoiceValue() == 1)
                {
                    return;
                }
                if (txtDriverMobile.Text != "")
                {
                    if (txtDriverMobile.Text.Length != 10)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Driver Mobile Must be 10 Digit!');", true);
                        return;
                    }
                }

                if (txtInvoiceNo1.Text == "" && txtInvoiceNo2.Text == "" && txtInvoiceNo3.Text == "" &&
     txtInvoiceNo4.Text == "" && txtInvoiceNo5.Text == "" &&
     txtInvoiceNo6.Text == "" && txtInvoiceNo7.Text == "" && txtInvoiceNo8.Text == "" &&
     txtInvoiceNo9.Text == "" && txtInvoiceNo10.Text == "" &&
     txtInvoiceNo11.Text == "" && txtInvoiceNo12.Text == "" && txtInvoiceNo13.Text == "" &&
     txtInvoiceNo14.Text == "" && txtInvoiceNo15.Text == "")
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                        "alert('You must have to Enter at least one Invoice No');", true);
                    return;
                }

                if (txtEwayBillNo1.Text == "" && txtEwayBillNo2.Text == "" && txtEwayBillNo3.Text == "" &&
     txtEwayBillNo4.Text == "" && txtEwayBillNo5.Text == "" &&
     txtEwayBillNo6.Text == "" && txtEwayBillNo7.Text == "" && txtEwayBillNo8.Text == "" &&
     txtEwayBillNo9.Text == "" && txtEwayBillNo10.Text == "" &&
     txtEwayBillNo11.Text == "" && txtEwayBillNo12.Text == "" && txtEwayBillNo13.Text == "" &&
     txtEwayBillNo14.Text == "" && txtEwayBillNo15.Text == "")
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                        "alert('You must have to Enter at least one Eway Bill No');", true);
                    return;
                }

                if (txtDocumentNo1.Text == "" && txtDocumentNo2.Text == "" && txtDocumentNo3.Text == "" &&
     txtDocumentNo4.Text == "" && txtDocumentNo5.Text == "" &&
     txtDocumentNo6.Text == "" && txtDocumentNo7.Text == "" && txtDocumentNo8.Text == "" &&
     txtDocumentNo9.Text == "" && txtDocumentNo10.Text == "" &&
     txtDocumentNo11.Text == "" && txtDocumentNo12.Text == "" && txtDocumentNo13.Text == "" &&
     txtDocumentNo14.Text == "" && txtDocumentNo15.Text == "")
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                        "alert('You must have to Enter at least one Document No');", true);
                    return;
                }
                /*
                if (txtEwayBillNo1.Text != "")
                {
                    if (txtEwayBillNo1.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }
                if (txtEwayBillNo2.Text != "")
                {
                    if (txtEwayBillNo2.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }
                if (txtEwayBillNo3.Text != "")
                {
                    if (txtEwayBillNo3.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }
                if (txtEwayBillNo4.Text != "")
                {
                    if (txtEwayBillNo4.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }
                if (txtEwayBillNo5.Text != "")
                {
                    if (txtEwayBillNo5.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }
                if (txtEwayBillNo6.Text != "")
                {
                    if (txtEwayBillNo6.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo7.Text != "")
                {
                    if (txtEwayBillNo7.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo8.Text != "")
                {
                    if (txtEwayBillNo8.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo9.Text != "")
                {
                    if (txtEwayBillNo9.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo10.Text != "")
                {
                    if (txtEwayBillNo10.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo11.Text != "")
                {
                    if (txtEwayBillNo11.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo12.Text != "")
                {
                    if (txtEwayBillNo12.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo13.Text != "")
                {
                    if (txtEwayBillNo13.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                if (txtEwayBillNo14.Text != "")
                {
                    if (txtEwayBillNo14.Text.Length != 12)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                            "alert('E-Way Bill No Must be 12 Digit!');", true);
                        return;
                    }
                }

                //if (txtEwayBillNo15.Text != "")
                //{
                //    if (txtEwayBillNo15.Text.Length != 12)
                //    {
                //        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script",
                //            "alert('E-Way Bill No Must be 12 Digit!');", true);
                //        return;
                //    }
                //}*/

                if (txtInvoiceNo1.Text != "")
                {
                    if (InvoiceDate1.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue1.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }
                if (txtInvoiceNo2.Text != "")
                {
                    if (InvoiceDate2.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue2.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }
                if (txtInvoiceNo3.Text != "")
                {
                    if (InvoiceDate3.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue3.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }
                if (txtInvoiceNo4.Text != "")
                {
                    if (InvoiceDate4.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue4.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }
                if (txtInvoiceNo5.Text != "")
                {
                    if (InvoiceDate5.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue5.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }
                if (txtInvoiceNo6.Text != "")
                {
                    if (InvoiceDate6.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue6.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo7.Text != "")
                {
                    if (InvoiceDate7.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue7.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo8.Text != "")
                {
                    if (InvoiceDate8.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue8.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo9.Text != "")
                {
                    if (InvoiceDate9.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue9.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo10.Text != "")
                {
                    if (InvoiceDate10.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue10.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo11.Text != "")
                {
                    if (InvoiceDate11.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue11.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo12.Text != "")
                {
                    if (InvoiceDate12.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue12.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo13.Text != "")
                {
                    if (InvoiceDate13.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue13.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo14.Text != "")
                {
                    if (InvoiceDate14.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue14.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtInvoiceNo15.Text != "")
                {
                    if (InvoiceDate15.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                        return;
                    }
                    //if (InvoiceValue15.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                    //    return;
                    //}
                }

                if (txtEwayBillNo1.Text != "")
                {
                    //if (txtValidUptoDate1.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate1.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }
                if (txtEwayBillNo2.Text != "")
                {
                    //if (txtValidUptoDate2.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate2.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }
                if (txtEwayBillNo3.Text != "")
                {
                    //if (txtValidUptoDate3.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate3.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }
                if (txtEwayBillNo4.Text != "")
                {
                    //if (txtValidUptoDate4.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate4.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }
                if (txtEwayBillNo5.Text != "")
                {
                    //if (txtValidUptoDate5.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate5.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }
                if (txtEwayBillNo6.Text != "")
                {
                    //if (txtValidUptoDate6.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate6.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo7.Text != "")
                {
                    //if (txtValidUptoDate7.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate7.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo8.Text != "")
                {
                    //if (txtValidUptoDate8.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate8.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo9.Text != "")
                {
                    //if (txtValidUptoDate9.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate9.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo10.Text != "")
                {
                    //if (txtValidUptoDate10.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate10.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo11.Text != "")
                {
                    //if (txtValidUptoDate11.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate11.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo12.Text != "")
                {
                    //if (txtValidUptoDate12.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate12.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo13.Text != "")
                {
                    //if (txtValidUptoDate13.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate13.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo14.Text != "")
                {
                    //if (txtValidUptoDate14.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate14.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }

                if (txtEwayBillNo15.Text != "")
                {
                    //if (txtValidUptoDate15.Text == "")
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                    //    return;
                    //}
                    if (txtEwayBillDate15.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                        return;
                    }
                }


                if (ddlMode.SelectedIndex == 0)
                {
                    Count = GetLRCount();
                }
                else if (ddlMode.SelectedIndex == 1)
                {
                    Count = GetMnaualLRCount();
                }
                else
                {
                    Count = GetLRCount();
                    if (Count == 0)
                    {
                        Count = GetMnaualLRCount();
                    }
                }
                Count1 = GetLCCount();

                if (HFid.Value == "")
                {
                    if (CheckDuplicateLR() > 0)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LR No!');", true);
                    }
                    else if (Count == 0)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for Consignment No!');TxtLrNo.focus;", true);
                    }
                    //else if (CheckDuplicateLC() > 0)
                    //{
                    //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LC No!');", true);
                    //}
                    else if (Count1 == 0 && txtLCNo.Text != "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for LC No!');txtLCNo.focus;", true);
                    }
                    else
                    {
                        if (Session["GroupSessionId"] == null)
                        {
                            Session["GroupSessionId"] = NewGroupId();
                        }

                        ParameterName = new ArrayList();
                        ParameterValue = new ArrayList();
                        spName = "Proc_BookLR";
                        ParameterName.Add("@Action"); ParameterValue.Add("InsertTMP");
                        ParameterName.Add("@BookType"); ParameterValue.Add(ddVType.SelectedValue);
                        ParameterName.Add("@BookLrNo"); ParameterValue.Add(TxtLrNo.Text);
                        ParameterName.Add("@BookDate"); ParameterValue.Add(txtdate.Text);
                        ParameterName.Add("@SourceFromid"); ParameterValue.Add(ddSourceFrom.SelectedValue);
                        ParameterName.Add("@SourceToid"); ParameterValue.Add(ddSourceTo.SelectedValue);
                        ParameterName.Add("@BookDestination"); ParameterValue.Add(txtDestination.Text);
                        ParameterName.Add("@BookConsiginorid"); ParameterValue.Add(ddConsignor.SelectedValue);
                        ParameterName.Add("@BookConsigieeid"); ParameterValue.Add(ddConsignee.SelectedValue);
                        ParameterName.Add("@GSTPay"); ParameterValue.Add(ddlGSTPay.SelectedValue);
                        ParameterName.Add("@BookPartyid"); ParameterValue.Add(0);
                        ParameterName.Add("@BookVchid"); ParameterValue.Add(ddlVchNo.SelectedValue);
                        ParameterName.Add("@BookInvNo"); ParameterValue.Add(txtinvno.Text);
                        ParameterName.Add("@BookInvDate"); ParameterValue.Add(txtinvdate.Text);
                        ParameterName.Add("@BookAmount"); ParameterValue.Add(txtAmount.Text);
                        ParameterName.Add("@BookEwayBillNo"); ParameterValue.Add(txtEwayBillNo.Text);
                        ParameterName.Add("@BookFromDate"); ParameterValue.Add(txtfromdate.Text);
                        ParameterName.Add("@BookValidDate"); ParameterValue.Add(txttodate.Text);
                        ParameterName.Add("@BookShipNo"); ParameterValue.Add(txtShipNo.Text);
                        ParameterName.Add("@BookDocNo"); ParameterValue.Add(txtDocNo.Text);
                        ParameterName.Add("@BookLoading"); ParameterValue.Add(txtLoading.Text);
                        ParameterName.Add("@BookRemarks"); ParameterValue.Add(txtRemarks.Text);
                        ParameterName.Add("@ProdName"); ParameterValue.Add(txtProduct.Text.Trim()); SaveProduct(txtProduct.Text.Trim());
                        ParameterName.Add("@UnitName"); ParameterValue.Add(ddlunit.SelectedValue);
                        ParameterName.Add("@BookQty"); ParameterValue.Add(txtQty.Text);
                        ParameterName.Add("@ActQty"); ParameterValue.Add(0);
                        ParameterName.Add("@BookBranchid"); ParameterValue.Add(hfbrid.Value);
                        ParameterName.Add("@BookDriverMobileNo"); ParameterValue.Add(txtDriverMobile.Text);
                        ParameterName.Add("@BookCustSize"); ParameterValue.Add(ddlSize.SelectedValue);
                        ParameterName.Add("@BookRemark1"); ParameterValue.Add(txtRemarks1.Text);
                        ParameterName.Add("@DriverName"); ParameterValue.Add(txtDriverName.Text);
                        ParameterName.Add("@SealNo"); ParameterValue.Add(txtsealno.Text);
                        ParameterName.Add("@BookBillTo"); ParameterValue.Add(ddlBillto.SelectedValue);
                        ParameterName.Add("@Mtype"); ParameterValue.Add(ddlMode.SelectedValue);
                        ParameterName.Add("@UserId"); ParameterValue.Add(Convert.ToInt32(Session["UserId"]));
                        ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                        ParameterName.Add("@DriverUID"); ParameterValue.Add(txtDriverAdhar.Text);
                        ParameterName.Add("@LCNo"); ParameterValue.Add(txtLCNo.Text);
                        IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                        if (IntResult > 0)
                        {
                            string MaxBookId1 = MaxBookId();
                            spName = "Proc_DocumentDetails";

                            if (txtDocumentNo1.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo1.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtDocumentNo2.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo2.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtDocumentNo3.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo3.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtDocumentNo4.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo4.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtDocumentNo5.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo5.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtDocumentNo6.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo6.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo7.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo7.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo8.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo8.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo9.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo9.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo10.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo10.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo11.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo11.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo12.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo12.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo13.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo13.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo14.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo14.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtDocumentNo15.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@DocumentNo"); ParameterValue.Add(txtDocumentNo15.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            spName = "Proc_InvoiceDetails";
                            if (txtInvoiceNo1.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo1.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate1.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue1.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit1.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtInvoiceNo2.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo2.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate2.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue2.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit2.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtInvoiceNo3.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo3.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate3.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue3.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit3.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtInvoiceNo4.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo4.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate4.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue4.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit4.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtInvoiceNo5.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo5.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate5.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue5.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit5.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtInvoiceNo6.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo6.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate6.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue6.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit6.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo7.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo7.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate7.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue7.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit7.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo8.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo8.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate8.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue8.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit8.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo9.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo9.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate9.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue9.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit9.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo10.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo10.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate10.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue10.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit10.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo11.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo11.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate11.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue11.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit11.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo12.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo12.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate12.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue12.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit12.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo13.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo13.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate13.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue13.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit13.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo14.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo14.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate14.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue14.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit14.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            if (txtInvoiceNo15.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@InvoiceNo"); ParameterValue.Add(txtInvoiceNo15.Text);
                                ParameterName.Add("@InvoiceDate"); ParameterValue.Add(InvoiceDate15.Text);
                                ParameterName.Add("@InvoiceValue"); ParameterValue.Add(InvoiceValue15.Text);
                                ParameterName.Add("@InvoiceUnit"); ParameterValue.Add(InvoiceUnit15.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            spName = "Proc_EwayBillDetails";
                            if (txtEwayBillNo1.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo1.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate1.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate1.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtEwayBillNo2.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo2.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate2.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate2.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtEwayBillNo3.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo3.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate3.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate3.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtEwayBillNo4.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo4.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate4.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate4.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            if (txtEwayBillNo5.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo5.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate5.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate5.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            // 6
                            if (txtEwayBillNo6.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo6.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate6.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate6.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 7
                            if (txtEwayBillNo7.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo7.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate7.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate7.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 8
                            if (txtEwayBillNo8.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo8.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate8.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate8.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 9
                            if (txtEwayBillNo9.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo9.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate9.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate9.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 10
                            if (txtEwayBillNo10.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo10.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate10.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate10.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 11
                            if (txtEwayBillNo11.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo11.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate11.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate11.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 12
                            if (txtEwayBillNo12.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo12.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate12.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate12.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 13
                            if (txtEwayBillNo13.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo13.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate13.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate13.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 14
                            if (txtEwayBillNo14.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo14.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate14.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate14.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }

                            // 15
                            if (txtEwayBillNo15.Text != "")
                            {
                                ParameterName = new ArrayList();
                                ParameterValue = new ArrayList();
                                ParameterName.Add("@Action"); ParameterValue.Add("InsertTmp");
                                ParameterName.Add("@LrNo"); ParameterValue.Add(TxtLrNo.Text);
                                ParameterName.Add("@EwayBillNo"); ParameterValue.Add(txtEwayBillNo15.Text);
                                ParameterName.Add("@BillDate"); ParameterValue.Add(txtEwayBillDate15.Text);
                                ParameterName.Add("@BillUptoDate"); ParameterValue.Add(txtValidUptoDate15.Text);
                                ParameterName.Add("@GroupId"); ParameterValue.Add(Convert.ToInt64(Session["GroupSessionId"]));
                                IntResult = con.ExecuteQueryByStoreProcedureWithParameter(spName, ParameterName, ParameterValue);
                            }
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data SuccessFully inserted!');", true);
                            if (txtEwayBillNo.Text.Trim() != "")
                            {
                                Session["PrintBookId"] = MaxBookId();
                            }
                            ClearForMulti();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        protected void btnVerify_Click(object sender, EventArgs e)
        {
            try
            {
                string msg;
                if (ddlMode.SelectedIndex == 2)
                {
                    if (Session["GroupSessionId"] == null && HFid.Value == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to add atleast one LR');", true);
                        return;
                    }
                }

                if (ddlMode.SelectedIndex != 2)
                {
                    if (txtDriverMobile.Text != "")
                    {
                        if (txtDriverMobile.Text.Length != 10)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Driver Mobile Must be 10 Digit!');", true);
                            return;
                        }
                    }

                    // Invoice No check
                    if (txtInvoiceNo1.Text == "" && txtInvoiceNo2.Text == "" && txtInvoiceNo3.Text == "" && txtInvoiceNo4.Text == "" && txtInvoiceNo5.Text == "" &&
                        txtInvoiceNo6.Text == "" && txtInvoiceNo7.Text == "" && txtInvoiceNo8.Text == "" && txtInvoiceNo9.Text == "" && txtInvoiceNo10.Text == "" &&
                        txtInvoiceNo11.Text == "" && txtInvoiceNo12.Text == "" && txtInvoiceNo13.Text == "" && txtInvoiceNo14.Text == "" && txtInvoiceNo15.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter atlest one Invoice No');", true);
                        return;
                    }

                    // EwayBill No check
                    if (txtEwayBillNo1.Text == "" && txtEwayBillNo2.Text == "" && txtEwayBillNo3.Text == "" && txtEwayBillNo4.Text == "" && txtEwayBillNo5.Text == "" &&
                        txtEwayBillNo6.Text == "" && txtEwayBillNo7.Text == "" && txtEwayBillNo8.Text == "" && txtEwayBillNo9.Text == "" && txtEwayBillNo10.Text == "" &&
                        txtEwayBillNo11.Text == "" && txtEwayBillNo12.Text == "" && txtEwayBillNo13.Text == "" && txtEwayBillNo14.Text == "" && txtEwayBillNo15.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter atlest one LR No');", true);
                        return;
                    }

                    // Document No check
                    if (txtDocumentNo1.Text == "" && txtDocumentNo2.Text == "" && txtDocumentNo3.Text == "" && txtDocumentNo4.Text == "" && txtDocumentNo5.Text == "" &&
                        txtDocumentNo6.Text == "" && txtDocumentNo7.Text == "" && txtDocumentNo8.Text == "" && txtDocumentNo9.Text == "" && txtDocumentNo10.Text == "" &&
                        txtDocumentNo11.Text == "" && txtDocumentNo12.Text == "" && txtDocumentNo13.Text == "" && txtDocumentNo14.Text == "" && txtDocumentNo15.Text == "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter atlest one Document No');", true);
                        return;
                    }


                    if (txtInvoiceNo1.Text != "")
                    {
                        if (InvoiceDate1.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue1.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }
                    if (txtInvoiceNo2.Text != "")
                    {
                        if (InvoiceDate2.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue2.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }
                    if (txtInvoiceNo3.Text != "")
                    {
                        if (InvoiceDate3.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue3.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }
                    if (txtInvoiceNo4.Text != "")
                    {
                        if (InvoiceDate4.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue4.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }   
                    if (txtInvoiceNo5.Text != "")
                    {
                        if (InvoiceDate5.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue5.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }
                    // 6
                    if (txtInvoiceNo6.Text != "")
                    {
                        if (InvoiceDate6.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue6.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 7
                    if (txtInvoiceNo7.Text != "")
                    {
                        if (InvoiceDate7.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue7.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 8
                    if (txtInvoiceNo8.Text != "")
                    {
                        if (InvoiceDate8.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue8.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 9
                    if (txtInvoiceNo9.Text != "")
                    {
                        if (InvoiceDate9.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue9.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 10
                    if (txtInvoiceNo10.Text != "")
                    {
                        if (InvoiceDate10.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue10.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 11
                    if (txtInvoiceNo11.Text != "")
                    {
                        if (InvoiceDate11.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue11.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 12
                    if (txtInvoiceNo12.Text != "")
                    {
                        if (InvoiceDate12.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue12.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 13
                    if (txtInvoiceNo13.Text != "")
                    {
                        if (InvoiceDate13.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue13.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 14
                    if (txtInvoiceNo14.Text != "")
                    {
                        if (InvoiceDate14.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue14.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    // 15
                    if (txtInvoiceNo15.Text != "")
                    {
                        if (InvoiceDate15.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Date');", true);
                            return;
                        }
                        //if (InvoiceValue15.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Invoice Value');", true);
                        //    return;
                        //}
                    }

                    if (txtEwayBillNo1.Text != "")
                    {
                        //if (txtValidUptoDate1.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate1.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo2.Text != "")
                    {
                        //if (txtValidUptoDate2.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate2.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo3.Text != "")
                    {
                        //if (txtValidUptoDate3.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate3.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo4.Text != "")
                    {
                        //if (txtValidUptoDate4.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate4.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo5.Text != "")
                    {
                        //if (txtValidUptoDate5.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate5.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }
                    // 6
                    if (txtEwayBillNo6.Text != "")
                    {
                        //if (txtValidUptoDate6.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate6.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 7
                    if (txtEwayBillNo7.Text != "")
                    {
                        //if (txtValidUptoDate7.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate7.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 8
                    if (txtEwayBillNo8.Text != "")
                    {
                        //if (txtValidUptoDate8.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate8.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 9
                    if (txtEwayBillNo9.Text != "")
                    {
                        //if (txtValidUptoDate9.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate9.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 10
                    if (txtEwayBillNo10.Text != "")
                    {
                        //if (txtValidUptoDate10.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate10.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 11
                    if (txtEwayBillNo11.Text != "")
                    {
                        //if (txtValidUptoDate11.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate11.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 12
                    if (txtEwayBillNo12.Text != "")
                    {
                        //if (txtValidUptoDate12.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate12.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 13
                    if (txtEwayBillNo13.Text != "")
                    {
                        //if (txtValidUptoDate13.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate13.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 14
                    if (txtEwayBillNo14.Text != "")
                    {
                        //if (txtValidUptoDate14.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate14.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }

                    // 15
                    if (txtEwayBillNo15.Text != "")
                    {
                        //if (txtValidUptoDate15.Text == "")
                        //{
                        //    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Valid Upto Date');", true);
                        //    return;
                        //}
                        if (txtEwayBillDate15.Text == "")
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('You must have to Enter Eway Bill Date');", true);
                            return;
                        }
                    }


                    int Count = 0;
                    int Count1 = 0;
                    if (ddlMode.SelectedIndex == 0)
                    {
                        Count = GetLRCount();
                    }
                    else if (ddlMode.SelectedIndex == 1)
                    {
                        Count = GetMnaualLRCount();
                    }
                    else
                    {
                        Count = GetLRCount();
                        if (Count == 0)
                        {
                            Count = GetMnaualLRCount();
                        }
                    }

                    Count1 = GetLCCount();

                    if (CheckInvoiceValue() == 1)
                    {
                        return;
                    }

                    if (HFid.Value == "")
                    {
                        if (CheckDuplicateLR() > 0)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LR No!');", true);
                            return;
                        }
                    }
                    else
                    {
                        if (CheckDuplicateLR1() > 0)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LR No!');", true);
                            return;
                        }
                    }

                    /*
                    if (HFid.Value == "")
                    {
                        if (CheckDuplicateLC() > 0)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LC No!');", true);
                            return;
                        }
                    }
                    else
                    {
                        if (CheckDuplicateLC1() > 0)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Data Already Exits on this LC No!');", true);
                            return;
                        }
                    }*/


                    if (Count == 0)
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for Consignment No!');TxtLrNo.focus;", true);
                        return;
                    }

                    if (Count1 == 0 && txtLCNo.Text != "")
                    {
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('please contact to admin for lc no!');txtlcno.focus;", true);
                        return;
                    }
                }
                /*
                if (ddlMode.SelectedIndex != 2)
                {
                    if (txtEwayBillNo1.Text != "")
                    {
                        if (txtEwayBillNo1.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo2.Text != "")
                    {
                        if (txtEwayBillNo2.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo3.Text != "")
                    {
                        if (txtEwayBillNo3.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo4.Text != "")
                    {
                        if (txtEwayBillNo4.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                    if (txtEwayBillNo5.Text != "")
                    {
                        if (txtEwayBillNo5.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                    // 6
                    if (txtEwayBillNo6.Text != "")
                    {
                        if (txtEwayBillNo6.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 7
                    if (txtEwayBillNo7.Text != "")
                    {
                        if (txtEwayBillNo7.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 8
                    if (txtEwayBillNo8.Text != "")
                    {
                        if (txtEwayBillNo8.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 9
                    if (txtEwayBillNo9.Text != "")
                    {
                        if (txtEwayBillNo9.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 10
                    if (txtEwayBillNo10.Text != "")
                    {
                        if (txtEwayBillNo10.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 11
                    if (txtEwayBillNo11.Text != "")
                    {
                        if (txtEwayBillNo11.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 12
                    if (txtEwayBillNo12.Text != "")
                    {
                        if (txtEwayBillNo12.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 13
                    if (txtEwayBillNo13.Text != "")
                    {
                        if (txtEwayBillNo13.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 14
                    if (txtEwayBillNo14.Text != "")
                    {
                        if (txtEwayBillNo14.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }

                    // 15
                    if (txtEwayBillNo15.Text != "")
                    {
                        if (txtEwayBillNo15.Text.Length != 12)
                        {
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('E-Way Bill No Must be 12 Digit!');", true);
                            return;
                        }
                    }
                }*/

                if ((cm.GetUserDetails().Rows[0]["Mobile"]).ToString() == "")
                {
                    msg = "alert('Please Input User No and again login.');";
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", msg, true);
                    return;
                }

                string OTPNumb = cm.FiveDigits();
                if (Convert.ToInt32((cm.GetCompanyData()).Rows[0]["isSMSServiceEnabled"]) == 0)
                {
                    Session["OTP"] = OTPNumb;
                    txtCode.Text = Session["OTP"].ToString();
                }
                else
                {
                    string MsgRes = cm.SendOTP(OTPNumb, (cm.GetUserDetails().Rows[0]["Mobile"]).ToString());
                    if (MsgRes == "SUCCESS" || ((MsgRes != "SUCCESS") && (Session["UserType"].ToString() == "A")))
                    {
                        msg = "alert('OTP Sent.Please check your Mobile');";
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", msg, true);
                        Session["OTP"] = OTPNumb;
                    }
                    else
                    {
                        msg = "alert('Send OTP Failed');";
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", msg, true);
                        Session["OTP"] = OTPNumb;
                    }
                }
                divsave.Visible = true;
                btnSaveLocal.Visible = false;
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP: | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        protected void ddConsignor_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddConsignor.SelectedIndex > 0)
            {
                ParameterName = new ArrayList();
                ParameterValue = new ArrayList();
                string spName = "Proc_Consign";
                try
                {
                    ParameterName.Add("@Action"); ParameterValue.Add("Get");
                    ParameterName.Add("@Id"); ParameterValue.Add(ddConsignor.SelectedValue);
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        lblConsignorAddress.Text = dt.Rows[0]["ConsgAdd"].ToString();
                        lblConsignorGST.Text = dt.Rows[0]["ConsgGSTNo"].ToString();
                        lblConsignorState.Text = dt.Rows[0]["ConsState"].ToString();
                        ddSourceFrom.SelectedValue = dt.Rows[0]["ConsSource"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
                }
            }
        }
        protected void ddlVchNo_SelectedIndexChanged(object sender, EventArgs e)
        {

            ParameterName = new ArrayList();
            ParameterValue = new ArrayList();
            string spName = "Proc_Vehicle";
            try
            {
                ParameterName.Add("@Action"); ParameterValue.Add("Get");
                ParameterName.Add("@Id"); ParameterValue.Add(ddlVchNo.SelectedValue);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    lblVchSize.Text = dt.Rows[0]["Size"].ToString();
                }

                ParameterName = new ArrayList();
                ParameterValue = new ArrayList();
                spName = "Proc_VajraLock";
                ParameterName.Add("@Action"); ParameterValue.Add("GetByVehicleNo");
                ParameterName.Add("@VehicleNo"); ParameterValue.Add(ddlVchNo.SelectedItem.Text);
                ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                dt = ds.Tables[0];
                if (dt.Rows.Count > 0)
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Vajra Lock Already Installed In This Vehicle');", true);
                }
                else
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Vajra Lock is not installed in this vehicle');", true);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }

        }
        protected void TextBox1_TextChanged(object sender, EventArgs e)
        {
            if (TextBox1.Text.Trim() != "")
            {
                ParameterName = new ArrayList();
                ParameterValue = new ArrayList();
                string spName = "Proc_Vehicle";
                try
                {
                    ParameterName.Add("@Action"); ParameterValue.Add("ListByVehicleNo");
                    ParameterName.Add("@VehicleNo"); ParameterValue.Add(TextBox1.Text.Trim());
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        ddlVchNo.DataSource = dt;
                        ddlVchNo.DataValueField = "Id";
                        ddlVchNo.DataTextField = "VehicleNo";
                        ddlVchNo.DataBind();
                        ddlVchNo.Items.Insert(0, new ListItem("--Select Vehicle--", "0"));

                        lblMessage.Text = "Total Result " + dt.Rows.Count.ToString();
                    }
                }
                catch (Exception ex)
                {
                    cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:" + spName + " | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
                }
            }
        }
        protected void TxtLrNo_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int Count = 0;

                if (ddlMode.SelectedIndex == 0)
                {
                    Count = GetLRCount();
                }
                else if (ddlMode.SelectedIndex == 1)
                {
                    Count = GetMnaualLRCount();
                }
                else
                {
                    Count = GetLRCount();
                    if (Count == 0)
                    {
                        Count = GetMnaualLRCount();
                    }
                }

                if (Count == 0)
                {
                    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('Please Contact To Admin for Consignment No!');TxtLrNo.focus;", true);
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }
        protected void ddlMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //if (ddlMode.SelectedIndex == 0 || ddlMode.SelectedIndex == 2)
                if (ddlMode.SelectedIndex == 0)
                {
                    txtdate.Enabled = false;
                    txtdate.Text = cm.GetCurrentDate();
                }
                else
                {
                    txtdate.Enabled = true;
                }

                if (ddlMode.SelectedIndex == 2)
                {
                    btnSaveLocal.Visible = true;
                }
                else
                {
                    btnSaveLocal.Visible = false;
                }
            }
            catch (Exception ex)
            {
                cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
            }
        }

        /*
        protected void txtLCNo_TextChanged(object sender, EventArgs e)
        {
            if(GetLCCount() == 0)
            {
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "script", "alert('please contact to admin for lc no!');txtlcno.focus;", true);
            }
        }
        */
        protected void ddConsignee_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ddConsignee.SelectedIndex > 0)
            {
                ParameterName = new ArrayList();
                ParameterValue = new ArrayList();
                string spName = "Proc_Consign";
                try
                {
                    ParameterName.Add("@Action"); ParameterValue.Add("Get");
                    ParameterName.Add("@Id"); ParameterValue.Add(ddConsignee.SelectedValue);
                    ds = con.GetDataSetFromSP(spName, ParameterName, ParameterValue);
                    dt = ds.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        lblConsigneeAddress.Text = dt.Rows[0]["ConsgAdd"].ToString();
                        lblConsigneeGst.Text = dt.Rows[0]["ConsgGSTNo"].ToString();
                        lblConsigneeState.Text = dt.Rows[0]["ConsState"].ToString();
                        ddSourceTo.SelectedValue = dt.Rows[0]["ConsSource"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    cm.WriteLog(DateTime.Now.ToString() + " | Page:" + Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath) + " | Db:" + Session["Database_"] + " | Method/Event:" + System.Reflection.MethodBase.GetCurrentMethod().Name + " | SP:  | User:" + Session["UserName"].ToString() + " | Err:" + ex.Message.ToString());
                }
            }
        }
        #endregion
    }
}
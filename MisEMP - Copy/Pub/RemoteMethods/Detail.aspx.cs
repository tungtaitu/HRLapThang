using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Threading;
using Pcc.Utils;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

public partial class Pub_RemoteMethods_OverTimeDetail : System.Web.UI.Page
{
    PersonnelVN PL;
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getOverTimeDetail":
                string MID = GetMID(strEventName, Commfun.CheckParams("vou_no", Page));
                getOverTimeDetail(MID);
                break;
        }
    }

    #region " 取得加班單明細 "
    private void getOverTimeDetail(string MID)
    {
        PL = new PersonnelVN();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            DataTable dt = PL.Get_OverTimeD_List(MID);
            get_OverTimeDetail.DataSource = dt;
            get_OverTimeDetail.DataBind();
        }
        catch (Exception ex)
        {
            sbErrMsg.Append("取得資料發生錯誤！！");
            sbErrMsg.Append(Commfun.JSStringFormat(ex.Message));
        }

        //輸出結果
        if (sbErrMsg.ToString().Trim() == "") //OK
        {

            System.IO.StringWriter sw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htw = new HtmlTextWriter(sw);
            HtmlForm hf = new HtmlForm();
            Controls.Add(hf);
            hf.Controls.Add(litSpliter1);
            hf.Controls.Add(get_OverTimeDetail);
            hf.Controls.Add(litSpliter2);
            hf.RenderControl(htw);
            string strBody = sw.ToString();
            strBody = Regex.Split(strBody, "<#BreakChar#>", RegexOptions.IgnoreCase)[1];
            strBody = Commfun.JSStringFormat(strBody);

            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:true");
            sbServerMsg.Append(",ServerMsg:'OK！'");
            sbServerMsg.Append(",Result:'" + strBody + "'");
            sbServerMsg.Append("}");
        }
        else //execute faile
        {
            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:false");
            sbServerMsg.Append(",ServerMsg:'" + Commfun.JSStringFormat(sbErrMsg.ToString().Trim()) + "'");
            sbServerMsg.Append("}");
        }
        Thread.Sleep(100);
        Response.Clear();
        Response.Write(sbServerMsg.ToString());
        Response.End();
    }
    protected void get_OverTimeDetail_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Cells[0].Text = Convert.ToString(e.Row.RowIndex + 1);
            string start_date = ((Label)e.Row.FindControl("lblStart_dateD")).Text.Trim();
            string end_date = ((Label)e.Row.FindControl("lblEnd_dateD")).Text.Trim();
            ((Label)e.Row.FindControl("lblStart_dateD")).Text = DateTimeUtils.ConvertString2FormatDatetime(start_date, "yyyy/MM/dd HH:mm");
            ((Label)e.Row.FindControl("lblEnd_dateD")).Text = DateTimeUtils.ConvertString2FormatDatetime(end_date, "yyyy/MM/dd HH:mm");

        }
    }
    #endregion

    private string GetMID(string strEventName,string vou_no)
    {
        PL = new PersonnelVN();
        string MID = "";
        switch (strEventName)
        {
            case "getOverTimeDetail":
                ArrayList AL = new ArrayList();
                AL.Add("vou_no : " + vou_no);
                DataTable dt = PL.SimpleGetData("overtimem", AL, "");
                MID = Commfun.CheckDBNull(dt.Rows[0]["ID"]);
                break;
        }
        return MID;
    }
    protected override void InitializeCulture()
    {
        // 用Session來存儲存言信息
        if (Session["PreferredCulture"] == null)
            Session["PreferredCulture"] = Request.UserLanguages[0];
        string UserCulture = Session["PreferredCulture"].ToString();
        if (UserCulture != "")
        {
            //根據Session的值重新綁定語言代碼
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(UserCulture);
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(UserCulture);
        }
    }
}

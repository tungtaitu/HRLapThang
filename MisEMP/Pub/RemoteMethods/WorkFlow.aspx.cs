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

public partial class Pub_RemoteMethods_WorkFlow : System.Web.UI.Page
{
    PersonnelVN PL;
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getDetail":
                string flow_id =  Commfun.CheckParams("FlowID", Page);
                getFlowUser(flow_id);
                break;
        }
    }
    private void getFlowUser(string flow_id)
    {
        PL = new PersonnelVN();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            DataTable dt = PL.Get_FlowUser(flow_id);
            gv_flowuser.DataSource = dt;
            gv_flowuser.DataBind();
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
            hf.Controls.Add(gv_flowuser);
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
    protected void gv_flowuser_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Cells[0].Text = Convert.ToString(e.Row.RowIndex + 1);
        }
    }


    protected override void InitializeCulture()
    {
        // 用Session来存储语言信息
        if (Session["PreferredCulture"] == null)
            Session["PreferredCulture"] = Request.UserLanguages[0];
        string UserCulture = Session["PreferredCulture"].ToString();
        if (UserCulture != "")
        {
            //根据Session的值重新绑定语言代码
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(UserCulture);
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(UserCulture);
        }
    }
}

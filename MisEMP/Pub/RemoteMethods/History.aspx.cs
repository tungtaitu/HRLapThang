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

public partial class Pub_RemoteMethods_History : BasePage
{
    PersonnelVN PL;
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getHistory":
                getHistory();
                break;
        }
    }
    private void getHistory()
    {
        PL = new PersonnelVN();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            DataTable dt = new DataTable();
            string vou_no = Commfun.CheckParams("vou_no", Page);
            string submit_id = Commfun.CheckParams("submit_id", Page);

            if (submit_id == "")
            {
                dt = PL.Get_History(Commfun.CheckParams("vou_no", Page));
            }
            else
            {
                BPMS bpms = new BPMS();
                dt = bpms.getBPMSSignData(vou_no,submit_id);
            }
            gv_history.DataSource = dt;
            gv_history.DataBind();
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
            hf.Controls.Add(gv_history);
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

    protected void gv_history_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string sign_mk = gv_history.DataKeys[e.Row.RowIndex].Value.ToString();
            if (sign_mk == "B")
            {
                e.Row.Cells[0].Text += "(授)";
            }
            else if (sign_mk == "A")
            {
                e.Row.Cells[0].Text += "(代)";
            }

            e.Row.Cells[1].Text = DateTimeUtils.ConvertString2FormatDatetime(e.Row.Cells[1].Text, "yyyy/MM/dd HH:mm:ss");

            string action = e.Row.Cells[3].Text;
            if (action == "申請")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/Apply.gif'>";
            }
            else if (action == "通知修改")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyInform.gif'>";
            }
            else if (action == "重新申請")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyRe.gif'>";
            }
            else if (action == "錯誤更正")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyRecover.gif'>";
            }
            else if (action == "簽核")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplySign.gif'>";
            }
            else if (action == "核准")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyCheck.gif'>";
            }
            else if (action == "撤單")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyWithdraw.gif'>";
            }
            else if (action == "駁回")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyReject.gif'>";
            }
            else if (action == "退單")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyReturn.gif'>";
            }
            else if (action == "強制退審")
            {
                e.Row.Cells[3].Text = "<img src='../../App_Themes/pwfBody/images/SignImg/ApplyCallBack.gif'>";
            }
        }
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


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
using PccCommonForC;

public partial class Pub_RemoteMethods_Detail_Over : System.Web.UI.Page
{
    PersonnelVN PL;
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getDetail_Over":             
                  string MID = GetMID(strEventName, Commfun.CheckParams("vou_no", Page));
                getDetail_Over(MID);
                break;
        }
    }

    #region " Lay Detail Over" 
    private void getDetail_Over(string MID)
    {       
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            DataTable dt = GetAppOvertimed(MID);
            get_Detail.DataSource = dt;
            get_Detail.DataBind();
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
            hf.Controls.Add(get_Detail);
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

    private DataTable GetAppOvertimed(string m_id)
    {
        PccMsg myMsg = new PccMsg();
        myMsg.CreateFirstNode("m_id", m_id);
        db_OverApplyGroup1 db = new db_OverApplyGroup1(ConfigurationSettings.AppSettings["PLVNConnectionType"], ConfigurationSettings.AppSettings["PLVNConnectionServer"], ConfigurationSettings.AppSettings["PLVNConnectionDB"], ConfigurationSettings.AppSettings["PLVNConnectionUser"], ConfigurationSettings.AppSettings["PLVNConnectionPwd"]);
        return db.pro_get_AppOvertimeddByM_ID(myMsg.GetXmlStr).Tables[0];

    }

    private string GetMID(string strEventName, string vou_no)
    {       
        string MID = "";
        switch (strEventName)
        {
            case "getDetail_Over":
                ArrayList AL = new ArrayList();
                AL.Add("vou_no : " + vou_no);
                DataTable dt = getDetailOver(vou_no);
                MID = Commfun.CheckDBNull(dt.Rows[0]["m_id"]);
                break;
        }
        return MID;
    }

    public DataTable getDetailOver(string vou_no)
    {
        PersonnelVN plvn = new PersonnelVN();
        DataTable dt = plvn.DataBySqlStr("select m_id from AppOvertimem where vou_no = '" + vou_no + "'");
        return dt;
    }


    protected void get_Detail_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Cells[0].Text = Convert.ToString(e.Row.RowIndex + 1);
        }
    }
    #endregion
    
  
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

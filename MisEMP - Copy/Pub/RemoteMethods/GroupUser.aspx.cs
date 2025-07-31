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
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

public partial class Pub_RemoteMethods_GroupUser : System.Web.UI.Page
{
    PersonnelVN PLVN;
    protected void Page_Load(object sender, EventArgs e)
    {
        
          string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getDetail":
                string GroupID = Commfun.CheckParams("GroupID", Page);
                Get_GroupUser(GroupID);
                break;
        }
    }

    private void Get_GroupUser(string GroupID)
    {
        PLVN = new PersonnelVN();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            //string group_no = Get_GroupNO(GroupID);
            ArrayList AL = new ArrayList();
            AL.Add("group_no :" + Get_GroupNO(GroupID));
            
            DataTable dt = PLVN.SimpleGetData("vie_employee", AL, "");
            gv_GroupUser.DataSource = dt;
            gv_GroupUser.DataBind();
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
            hf.Controls.Add(gv_GroupUser);
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

    private string Get_GroupNO(string group_id)
    {
        PLVN = new PersonnelVN();

        string group_no = "";
        ArrayList AL = new ArrayList();
        AL.Add("ID:" + group_id);
        DataTable dt = PLVN.SimpleGetData("authgroup", AL, "");
        
        if (dt.Rows.Count > 0)
        {
            group_no = Commfun.CheckDBNull(dt.Rows[0]["group_no"]);
        }
        return group_no;
    }
}

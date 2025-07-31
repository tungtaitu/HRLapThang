using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Configuration;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Threading;
using System.Globalization;

public partial class Pub_RemoteMethods_WorkFlowSequence : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "getPowerDetail":
                getPowerDetail();
                break;
            case "delAuth":
                delAuth();
                break;
        }
    }

    #region 取得user詳細資料
    private void getPowerDetail()
    {
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            string strUserID = Convert.ToString(Request.Params["UserID"]);
            PersonnelVN plvn = new PersonnelVN();
            ArrayList al = new ArrayList();
            al.Add("owner_id:" + strUserID);
            DataTable dt = plvn.SimpleGetData("authority", al, "id");
            DataTable _childtb = CreatTmpDt().Clone();//creat child tb
            DataRow _childa = null;
            foreach (DataRow dr in dt.Select("", "owner_id"))
            {
                // get authority parent's data from vie_authorg
                al.Clear();
                _childa = _childtb.NewRow();
                //get target_unit、target_id
                _childa = getRowChild(dr["id"].ToString().Trim(),dr["target_unit"].ToString().Trim(), dr["target_id"].ToString().Trim(), dr["id"].ToString().Trim(),
                                            dr["owner_id"].ToString().Trim(), _childa);
                _childtb.Rows.Add(_childa);
            }
            gvPowerDetail.DataSource = _childtb;
            gvPowerDetail.DataBind();
        }
        catch(Exception ex)
        {
            sbErrMsg.Append("取得資料發生錯誤！！");
            sbErrMsg.Append(JSStringFormat(ex.Message));
        }

        //輸出結果
        if (sbErrMsg.ToString().Trim() == "") //OK
        {

            System.IO.StringWriter sw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htw = new HtmlTextWriter(sw);
            HtmlForm hf = new HtmlForm();
            Controls.Add(hf);
            hf.Controls.Add(litSpliter1);
            hf.Controls.Add(gvPowerDetail);
            hf.Controls.Add(litSpliter2);
            hf.RenderControl(htw);
            string strBody = sw.ToString();
            strBody = Regex.Split(strBody, "<#BreakChar#>", RegexOptions.IgnoreCase)[1];
            strBody = JSStringFormat(strBody);

            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:true");
            sbServerMsg.Append(",ServerMsg:'更新成功！！'");
            sbServerMsg.Append(",Result:'"+strBody+"'");
            sbServerMsg.Append("}");
        }
        else //execute faile
        {
            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:false");
            sbServerMsg.Append(",ServerMsg:'" + JSStringFormat(sbErrMsg.ToString().Trim()) + "'");
            sbServerMsg.Append("}");
        }
        Thread.Sleep(100);
        Response.Clear();
        Response.Write(sbServerMsg.ToString());
        Response.End();

     }

     private DataTable CreatTmpDt()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("id");
        dt.Columns.Add("comp_cname");
        dt.Columns.Add("comp_no");
        dt.Columns.Add("fact_cname");
        dt.Columns.Add("fact_no");
        dt.Columns.Add("dept_cname");
        dt.Columns.Add("dept_no");
        dt.Columns.Add("group_cname");
        dt.Columns.Add("group_no");
        dt.Columns.Add("authorityId");//authority id
        dt.Columns.Add("user_id");

        return dt;
    }

    private DataRow getRowChild(string id,string Kind, string belongId, string authorityId, string user_id, DataRow _da)
    {
        string _restr = "", _tbnm = "vie_authorg", _orderby = "comp_id", _al = "";
        string comp_cname = "", comp_no = "", fact_cname = "", fact_no = "", dept_cname = "", dept_no = "", group_cname = "", group_no = "";
        ArrayList al = new ArrayList();
        PersonnelVN plvn = new PersonnelVN();
        #region 'select  Kind'
        switch (Kind.Trim())
        {
            case "C":
                _al = "comp_id:";

                break;

            case "F":
                _al = "fact_id:";
                break;

            case "D":
                _al = "dept_id:";
                break;

            case "G":
                _al = "group_id:";
                break;
        }
        #endregion
        al.Add(_al + belongId);
        try
        {
            DataTable dt = plvn.SimpleGetData(_tbnm, al, _orderby);

            if (dt.Rows.Count != 0)
            {
                #region   'add data to _da'
                switch (Kind.Trim())
                {
                    case "C":
                        comp_cname = dt.Rows[0]["comp_cname"].ToString();
                        comp_no = dt.Rows[0]["comp_no"].ToString();
                        fact_cname = "全";
                        fact_no = "0";
                        dept_cname = "全";
                        dept_no = "0";
                        group_cname = "全";
                        group_no = "0";

                        break;

                    case "F":
                        comp_cname = dt.Rows[0]["comp_cname"].ToString();
                        comp_no = dt.Rows[0]["comp_no"].ToString();
                        fact_cname = dt.Rows[0]["fact_cname"].ToString();
                        fact_no = dt.Rows[0]["fact_no"].ToString();
                        dept_cname = "全";
                        dept_no = "0";
                        group_cname = "全";
                        group_no = "0";
                        break;

                    case "D":
                        comp_cname = dt.Rows[0]["comp_cname"].ToString();
                        comp_no = dt.Rows[0]["comp_no"].ToString();
                        fact_cname = dt.Rows[0]["fact_cname"].ToString();
                        fact_no = dt.Rows[0]["fact_no"].ToString();
                        dept_cname = dt.Rows[0]["dept_cname"].ToString();
                        dept_no = dt.Rows[0]["dept_no"].ToString();
                        group_cname = "全";
                        group_no = "0";
                        break;

                    case "G":
                        comp_cname = dt.Rows[0]["comp_cname"].ToString();
                        comp_no = dt.Rows[0]["comp_no"].ToString();
                        fact_cname = dt.Rows[0]["fact_cname"].ToString();
                        fact_no = dt.Rows[0]["fact_no"].ToString();
                        dept_cname = dt.Rows[0]["dept_cname"].ToString();
                        dept_no = dt.Rows[0]["dept_no"].ToString();
                        group_cname = dt.Rows[0]["group_cname"].ToString();
                        group_no = dt.Rows[0]["group_no"].ToString();
                        break;

                }
                #endregion

                _da["id"] = id;
                _da["comp_cname"] = comp_cname;
                _da["comp_no"] = comp_no;
                _da["fact_cname"] = fact_cname;
                _da["fact_no"] = fact_no;
                _da["dept_cname"] = dept_cname;
                _da["dept_no"] = dept_no;
                _da["group_cname"] = group_cname;
                _da["group_no"] = group_no;
                _da["authorityId"] = authorityId;
                _da["user_id"] = user_id;
            }

        }
        catch (Exception ex)
        {
            _restr = "error :" + ex.Message.Trim();// error message
        }


        return _da;
    }
    #endregion

    #region 刪除user權限
    private void delAuth()
    {
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        string strRecID = Convert.ToString(Request.Params["RecID"]);
        try
        {
            PersonnelVN plvn = new PersonnelVN();
            plvn.Delete_authority(strRecID);
        }
        catch (Exception ex)
        {
            sbErrMsg.Append("刪除資料發生錯誤！！");
            sbErrMsg.Append(JSStringFormat(ex.Message));
        }

        if (sbErrMsg.ToString().Trim() == "")
        {
            getPowerDetail();
        }
        else
        {
            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:false");
            sbServerMsg.Append(",ServerMsg:'" + JSStringFormat(sbErrMsg.ToString().Trim()) + "'");
            sbServerMsg.Append("}");

            Response.Clear();
            Response.Write(sbServerMsg.ToString());
            Response.End();
        }

    }
    #endregion

    private string JSStringFormat(string s)
    {
        return s.Replace("\r", "\\r").Replace("\n", "\\n").Replace("'", "\\'").Replace("\"", "\\\"");
    }

    protected void img_delchild_Click(object sender, ImageClickEventArgs e)
    {
        ImageButton imgbt = sender as ImageButton;
        GridViewRow gvr = (GridViewRow)imgbt.Parent.Parent;
        PersonnelVN plvn = new PersonnelVN();
         
        try
        {
            plvn.Delete_authority(gvPowerDetail.DataKeys[gvr.RowIndex][1].ToString());
            Commfun.ShowMsg("資料刪除成功！", Page);
                 
        }
        catch (Exception ex)
        {
            Commfun.ShowMsg("刪除失敗：" + ex.Message + " ！請聯絡管理人員", Page);
        }
       
    }
    protected void gvPowerDetail_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            ((Label)e.Row.FindControl("Label1")).Text = Convert.ToString(e.Row.RowIndex + 1);
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

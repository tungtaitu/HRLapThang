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
using System.Text.RegularExpressions;
using Pcc.Utils;
using System.Threading;
using System.Globalization;
public partial class Pub_RemoteMethods_OverTimeRestEnd : System.Web.UI.Page
{
    PersonnelVN PL;
    protected void Page_Load(object sender, EventArgs e)
    {
        string strEventName = Convert.ToString(Request.Params["EventName"]);
        switch (strEventName)
        {
            case "OverTimeRestEnd":
                string emp_no = Commfun.CheckParams("emp_no", Page);
                string start_date = Commfun.CheckParams("start_date", Page);
                string end_date = Commfun.CheckParams("end_date", Page);
                getDetail(emp_no, start_date, end_date);
                break;
        }
    }
    #region " ���o�а������ "
    private DataTable GetRestData(string emp_No, string start_date, string end_date)
    {
        OracleFn ora = new OracleFn();
        DataTable dt = ora.OverTimeRestEnd_Detail(emp_No, start_date, end_date);

        return dt;
    }
    private void getDetail(string emp_no,string start_date, string end_date)
    {
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder("");
        System.Text.StringBuilder sbServerMsg = new System.Text.StringBuilder("");

        try
        {
            DataTable dt = GetRestData(emp_no, start_date, end_date);
            grvDetail.DataSource = dt;
            grvDetail.DataBind();
        }
        catch (Exception ex)
        {
            sbErrMsg.Append("���o��Ƶo�Ϳ��~�I�I");
            sbErrMsg.Append(Commfun.JSStringFormat(ex.Message));
        }

        //��X���G
        if (sbErrMsg.ToString().Trim() == "") //OK
        {

            System.IO.StringWriter sw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter htw = new HtmlTextWriter(sw);
            HtmlForm hf = new HtmlForm();
            Controls.Add(hf);
            hf.Controls.Add(litSpliter1);
            hf.Controls.Add(grvDetail);
            hf.Controls.Add(litSpliter2);
            hf.RenderControl(htw);
            string strBody = sw.ToString();
            strBody = Regex.Split(strBody, "<#BreakChar#>", RegexOptions.IgnoreCase)[1];
            strBody = Commfun.JSStringFormat(strBody);

            sbServerMsg.Append("{");
            sbServerMsg.Append("IsOK:true");
            sbServerMsg.Append(",ServerMsg:'OK�I'");
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
    private double ConvertDouble(string s)
    {
        double dRes = 0;
        bool b = double.TryParse(s, out dRes);

        return dRes;
    }
    private string ConvertDateTime(string s)
    {
        string date = s;
        if (s.Length == 8)
        {
            date = s.Substring(0, 4) + "/" + s.Substring(4, 2)+"/"+s.Substring(6,2);
        }
        if (s.Length == 4)
        {
            date = s.Substring(0, 2) + ":" + s.Substring(2, 2);
        }
        if (s.Length == 14)
        {
            date = s.Substring(0, 4) + "/" + s.Substring(4, 2) + "/" + s.Substring(6, 2);
            date = date + s.Substring(8, 2) + ":" + s.Substring(10, 2);
        }
        return date;
    }
    protected void getDetail_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        GridViewRow row = e.Row;
        if (row.RowType == DataControlRowType.Header)
        {
            row.Cells[2].ColumnSpan = 2;
            row.Cells[3].Visible = false;
        }
        if (row.RowType == DataControlRowType.DataRow)
        {
            DataRowView dr = (DataRowView)row.DataItem;
            
            string rest_hours = dr["rest_hours"].ToString().Trim();
            string old_rest = dr["old_rest"].ToString().Trim();
            double dnew_rest = ConvertDouble(rest_hours) - ConvertDouble(old_rest);
            string start_time = dr["start_time"].ToString().Trim();
            string end_time = dr["end_time"].ToString().Trim();

            e.Row.Cells[0].Text = Convert.ToString(e.Row.RowIndex + 1);
            row.Cells[2].Text = ConvertDateTime(row.Cells[2].Text.Trim());
            row.Cells[3].Text = ConvertDateTime(start_time) + " - " + ConvertDateTime(end_time);
            row.Cells[7].Text = dnew_rest.ToString();
        }
    }
    #endregion    
    protected override void InitializeCulture()
    {
        // ��Session�Ӧs�x�s���H��
        if (Session["PreferredCulture"] == null)
            Session["PreferredCulture"] = Request.UserLanguages[0];
        string UserCulture = Session["PreferredCulture"].ToString();
        if (UserCulture != "")
        {
            //�ھ�Session���ȭ��s�j�w�y���N�X
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(UserCulture);
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(UserCulture);
        }
    }
  
}

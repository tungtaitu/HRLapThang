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
using Pcc.Utils.CollectionUtils;
using System.Threading;
using System.Globalization;

/// <summary>
/// 挑選所有在vie_employee 有帳號的使用者
/// </summary>
public partial class SalaryReport_AuthGroupApplication_PikcAllEmps : BasePage
{
    #region Page_Load
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            ChangeLang();
            
            Gw_PickUser.DataSource = GetUserData();
            Gw_PickUser.DataBind();
        }

    }
    #endregion

    #region Function
    //GetUserData
    private DataTable GetUserData()
    {
        PersonnelVN plvn = new PersonnelVN();
        DataTable dt = plvn.Get_Have_UserID_Emp(NoAuthFEC1.CompSelectValue.Trim(), NoAuthFEC1.FactSelectValue.Trim(), 
            NoAuthFEC1.DptSelectValue.Trim(), NoAuthFEC1.GrpSelectValue.Trim(), txtNm.Text.Trim(),"0");
        return dt;
    }
   
    #endregion

    #region Event
    protected void Gw_PickUser_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        Gw_PickUser.PageIndex = e.NewPageIndex;
        Gw_PickUser.DataSource = GetUserData();
        Gw_PickUser.DataBind();
    }
    protected void Gw_PickUser_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string[] param = this.hidSelEmployee.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);

            foreach (string item in param)
            {
                if (item == ((HiddenField)e.Row.Cells[0].FindControl("HidID")).Value)
                {
                    ((CheckBox)e.Row.Cells[0].FindControl("chkUserID")).Checked = true;
                    break;
                }
            }
            //e.Row.Cells[0].Text = Convert.ToString(e.Row.DataItemIndex + 1);
            //int FlagChk = Session["chkTypeList"].ToString().IndexOf(Gw_PickUser.DataKeys[e.Row.RowIndex][0].ToString() + ":" + Gw_PickUser.DataKeys[e.Row.RowIndex][1].ToString());

            //if (FlagChk != -1)
            //{
            //    e.Row.Cells[5].Text = string.Format("<input type='checkbox' name='chkType' checked='checked' id={0} value={1}>", "chk_" + e.Row.RowIndex, Gw_PickUser.DataKeys[e.Row.RowIndex].Value + ":" + e.Row.Cells[1].Text);
            //}
            //else
            //{
            //    e.Row.Cells[5].Text = string.Format("<input type='checkbox' name='chkType' id={0} value={1}>", "chk_" + e.Row.RowIndex, Gw_PickUser.DataKeys[e.Row.RowIndex].Value + ":" + e.Row.Cells[1].Text);
            //}
            //hidCheckAllCount.Value = Convert.ToString(e.Row.RowIndex + 1);//取得目前page所有筆數
        }

    }
    //取得挑選User
    protected void btChoice_Click(object sender, EventArgs e)
    {
        //if (Commfun.CheckForm("chkType", this.Page) != "")
        //{
        //    string[] UserInfo = Commfun.CheckForm("chkType", this.Page).Split(',');// get ID
        //    string UserNm = "", UserId = "";

        //    for (int i = 0; i < UserInfo.Length; i++)
        //    {
        //        UserId += UserInfo[i].Split(':')[0] + ",";
        //        UserNm += UserInfo[i].Split(':')[1] + ",";
        //    }
        //    UserId = UserId.Substring(0, UserId.Length - 1);//去掉多餘的,
        //    UserNm = UserNm.Substring(0, UserNm.Length - 1);

        //    // RegisterStartupScript("New", "<script language=javascript>CloseWin();</script>");
        //    ClientScript.RegisterStartupScript(GetType(), "New", "<script language=javascript>GetChoiceUserIno('" + UserId + "$" + UserNm + " ');</script>");
        //}
        if (hidSelEmployee.Value != "")
        {
            ClientScript.RegisterStartupScript(GetType(), "New", "<script language=javascript>GetChoiceUserIno('" + hidSelEmployee.Value + "$" + hidSelAgent.Value + " ');</script>");
        }
        else
        {
            Commfun.ShowMsg_ScriptManager("請挑選人員！", Page);
        }
    }
    //查詢資料
    protected void btQuery_Click(object sender, EventArgs e)
    {
        Gw_PickUser.DataSource = GetUserData();
        Gw_PickUser.DataBind();
    }
 
    
    #endregion

    #region ' chkeck Event '

    //分頁挑選
    protected void chkUserID_CheckedChanged(object sender, EventArgs e)
    {
        CheckBox ckb = sender as CheckBox;

        if (ckb.Checked == true)
        {
            this.hidSelEmployee.Value += "," + ((HiddenField)ckb.Parent.FindControl("HidID")).Value;
            this.hidSelAgent.Value += "," + ((HiddenField)ckb.Parent.FindControl("HidNM")).Value;
            if (hidSelEmployee.Value.Substring(0, 1) == ",")
            {
                hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
                hidSelAgent.Value = hidSelAgent.Value.Substring(1, hidSelAgent.Value.Length - 1);
            }
        }
        else
        {
            string[] param = this.hidSelEmployee.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);
            string[] param1 = this.hidSelAgent.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);

            this.hidSelEmployee.Value = "";
            this.hidSelAgent.Value = "";

            foreach (string item in param)
            {
                if (item != ((HiddenField)ckb.Parent.FindControl("HidID")).Value)
                {
                    this.hidSelEmployee.Value += "," + item;
                }
            }
            foreach (string item in param1)
            {
                if (item != ((HiddenField)ckb.Parent.FindControl("HidNM")).Value)
                {
                    this.hidSelAgent.Value += "," + item;
                }
            }
            if (hidSelEmployee.Value != "")
            {
                hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
                hidSelAgent.Value = hidSelAgent.Value.Substring(1, hidSelAgent.Value.Length - 1);
            }
        }
    }

    //單頁全選
    protected void chkAll_CheckedChanged(object sender, EventArgs e)
    {
        if (Gw_PickUser.Rows.Count > 0)
        {
            if (((CheckBox)Gw_PickUser.HeaderRow.FindControl("chkAll")).Checked == true)
            {
                for (int i = 0; i < Gw_PickUser.Rows.Count; i++)
                {
                    ((CheckBox)Gw_PickUser.Rows[i].FindControl("chkUserID")).Checked = true;
                    this.hidSelEmployee.Value += "," + ((HiddenField)Gw_PickUser.Rows[i].FindControl("HidID")).Value;
                    this.hidSelAgent.Value += "," + ((HiddenField)Gw_PickUser.Rows[i].FindControl("HidNM")).Value;
                    if (hidSelEmployee.Value.Substring(0, 1) == ",")
                    {
                        hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
                        hidSelAgent.Value = hidSelAgent.Value.Substring(1, hidSelAgent.Value.Length - 1);
                    }
                }
            }
            else
            {
                this.hidSelEmployee.Value = "";
                this.hidSelAgent.Value = "";
                for (int i = 0; i < Gw_PickUser.Rows.Count; i++)
                {
                    ((CheckBox)Gw_PickUser.Rows[i].FindControl("chkUserID")).Checked = false;
                    string[] param = this.hidSelEmployee.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);
                    string[] param1 = this.hidSelAgent.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);

                    foreach (string item in param)
                    {
                        if (item != ((HiddenField)Gw_PickUser.Rows[i].FindControl("HidID")).Value)
                        {
                            this.hidSelEmployee.Value += "," + item;
                        }
                    }
                    foreach (string item in param1)
                    {
                        if (item != ((HiddenField)Gw_PickUser.Rows[i].FindControl("HidNM")).Value)
                        {
                            this.hidSelAgent.Value += "," + item;
                        }
                    }
                    if (hidSelEmployee.Value != "")
                    {
                        hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
                        hidSelAgent.Value = hidSelAgent.Value.Substring(1, hidSelAgent.Value.Length - 1);
                    }
                }
            }
        }
    }


    #endregion

    #region "語系轉換"
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

    private void ChangeLang()
    {
        btQuery.Text = Resources.Strings.btnQuery;
        lblemp_nm.Text = Resources.Strings.emp_name+"：" ;
        btChoice.Text = Resources.Strings.btnPick;
    }
    #endregion
 
}

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

public partial class Pub_BaseControls_PickEmployee : System.Web.UI.UserControl
{
    PersonnelVN PL;
    string strUserID = "";
    #region Public Property
    public string ShowFlowID
    {
        get
        {
            return hidFlowID.Value;
        }
        set
        {
            hidFlowID.Value = value;
        }
    }
    public string ShowSelEmployee
    {
        get
        {
            return hidSelEmployee.Value;
        }
       
    }
    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {
        strUserID = BasePage.UserId;
        if (!IsPostBack)
        {
            GetCompList();
            ddlComp.SelectedValue = "0";
            ddlFact.SelectedValue = "0";
            ddlDept.SelectedValue = "0";
            ddlGroup.SelectedValue = "0";
        }
        //GetBaseData();
       
    }
    private void GetBaseData()
    {
        DataTable dt = null;
        gv_employee.DataSource = dt;
        gv_employee.DataBind();       
        txtQuery_name.Text = "";
        hidSelEmployee.Value = "";
    }

    #region ' gv_employee Function && Event"

    private void GetEmployeeList()
    {
        PL = new PersonnelVN();
        HashMap hm = new HashMap();
        hm.add("@user_id", strUserID);
        hm.add("@flow_id", hidFlowID.Value);
        hm.add("@comp_no", ddlComp.SelectedValue);
        hm.add("@fact_no", ddlFact.SelectedValue);
        hm.add("@dept_no", ddlDept.SelectedValue);
        hm.add("@group_no", ddlGroup.SelectedValue);
        hm.add("@emp_name", txtQuery_name.Text.Trim());
        hm.add("@emp_str", hidSelEmployee.Value.Trim());

        DataTable dt = PL.GetDataByProc("pro_get_employee_notadd_overtime", hm);
        gv_employee.DataSource = dt;
        gv_employee.DataBind();
    }
    protected void gv_employee_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        gv_employee.PageIndex = e.NewPageIndex;
        GetEmployeeList();
    }

    #endregion

    #region " 下拉式 "
    /// <summary>
    /// 使用者權限(公司別)
    /// </summary>
    private void GetCompList()
    {
        PL = new PersonnelVN();
        HashMap hm = new HashMap();
        hm.add("@user_id", strUserID);
        hm.add("@target_unit", "C");//目前正要查詢的單位
        hm.Add("@upd_no", "");//上一層單位的no
        DataTable dt = PL.GetDataByProc("pro_get_authority_userid", hm);
        ddlComp.Items.Clear();
        if (dt.Rows.Count > 0)
        {
            ddlComp.DataSource = dt;
            ddlComp.DataTextField = "comp_cname";
            ddlComp.DataValueField = "comp_no";
            ddlComp.DataBind();
        }
        ddlComp.Items.Insert(0, new ListItem("---請選擇---", "0"));
        GetFactList();

    }
    /// <summary>
    /// 使用者權限(廠別)
    /// </summary>
    private void GetFactList()
    {
        PL = new PersonnelVN();
        HashMap hm = new HashMap();
        hm.add("@user_id", strUserID);
        hm.add("@target_unit", "F");
        hm.Add("@upd_no", ddlComp.SelectedValue);
        DataTable dt = PL.GetDataByProc("pro_get_authority_userid", hm);
        ddlFact.Items.Clear();
        if (dt.Rows.Count > 0)
        {
            ddlFact.DataSource = dt;
            ddlFact.DataTextField = "fact_cname";
            ddlFact.DataValueField = "fact_no";
            ddlFact.DataBind();
        }
        ddlFact.Items.Insert(0, new ListItem("---請選擇---", "0"));
        GetDeptList();

    }
    /// <summary>
    /// 使用者權限(部門別)
    /// </summary>
    private void GetDeptList()
    {
        PL = new PersonnelVN();
        HashMap hm = new HashMap();
        hm.add("@user_id", strUserID);
        hm.add("@target_unit", "D");
        hm.Add("@upd_no", ddlFact.SelectedValue);
        DataTable dt = PL.GetDataByProc("pro_get_authority_userid", hm);
        ddlDept.Items.Clear();
        if (dt.Rows.Count > 0)
        {
            ddlDept.DataSource = dt;
            ddlDept.DataTextField = "dept_cname";
            ddlDept.DataValueField = "dept_no";
            ddlDept.DataBind();
        }
        ddlDept.Items.Insert(0, new ListItem("---請選擇---", "0"));
        GetGroupList();
    }
    /// <summary>
    /// 使用者權限(組別)
    /// </summary>
    private void GetGroupList()
    {
        PL = new PersonnelVN();
        HashMap hm = new HashMap();
        hm.add("@user_id", strUserID);
        hm.add("@target_unit", "G");
        hm.Add("@upd_no", ddlDept.SelectedValue);
        DataTable dt = PL.GetDataByProc("pro_get_authority_userid", hm);
        ddlGroup.Items.Clear();
        if (dt.Rows.Count > 0)
        {
            ddlGroup.DataSource = dt;
            ddlGroup.DataTextField = "group_cname";
            ddlGroup.DataValueField = "group_no";
            ddlGroup.DataBind();
        }
        ddlGroup.Items.Insert(0, new ListItem("---請選擇---", "0"));
    }
    #endregion

    #region ' DDL Event '
    protected void ddlComp_SelectedIndexChanged(object sender, EventArgs e)
    {
        GetFactList();
    }
    protected void ddlFact_SelectedIndexChanged(object sender, EventArgs e)
    {
        GetDeptList();
    }
    protected void ddlDept_SelectedIndexChanged(object sender, EventArgs e)
    {
        GetGroupList();
    }
    #endregion

    #region '  分頁chk Event '
    protected void chkUserID_CheckedChanged(object sender, EventArgs e)
    {
        CheckBox ckb = sender as CheckBox;

        if (ckb.Checked == true)
        {
            this.hidSelEmployee.Value += "," + ((HiddenField)ckb.Parent.FindControl("HidID")).Value;
            if (hidSelEmployee.Value.Substring(0, 1) == ",")
            {
                hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
            }
        }
        else
        {
            string[] param = this.hidSelEmployee.Value.Split(new char[] { ',' } , StringSplitOptions.RemoveEmptyEntries);

            this.hidSelEmployee.Value = "";

            foreach (string item in param)
            {
                if (item != ((HiddenField)ckb.Parent.FindControl("HidID")).Value)
                {
                    this.hidSelEmployee.Value += "," + item;
                }
            }
            if (hidSelEmployee.Value != "")
            {
                hidSelEmployee.Value = hidSelEmployee.Value.Substring(1, hidSelEmployee.Value.Length - 1);
            }
        }
    }
    #endregion

    #region " button "
   
    //查詢
    protected void btnQuery_name_Click(object sender, EventArgs e)
    {
        GetEmployeeList();
    }
    //清空查詢
    protected void btnSelClear_Click(object sender, EventArgs e)
    {
        ddlComp.SelectedValue = "0";
        ddlFact.SelectedValue = "0";
        ddlDept.SelectedValue = "0";
        ddlGroup.SelectedValue = "0";
        txtQuery_name.Text = "";
    }
    #endregion
}

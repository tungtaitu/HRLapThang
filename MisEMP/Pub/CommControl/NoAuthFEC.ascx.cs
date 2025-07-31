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
/// <summary>
/// FindEmpCondition
/// 查詢員工的基本Query Conditions
/// 公司別、廠別、部門、組別
/// 無權限
/// </summary>
public partial class Pub_CommControl_NoAuthFEC : System.Web.UI.UserControl 
{
    #region ' Variables '
    protected bool _ddlgroupVis = true;
    protected bool _ddlDptVis = true;
    protected bool _ddlfactVis = true;//預設顯示
    #endregion

    #region 'Public property value'

    #region  ' select value '
    /// <summary>
    /// 取的公司別
    /// value:comp_id
    /// </summary>
    public string CompSelectValue
    {
        get
        {  
            return ddl_comp.SelectedValue.ToString();
        }
        set
        {
            ddl_comp.SelectedValue = value;
        }
    }
    /// <summary>
    /// 廠別
    /// value:fact_id
    /// </summary>
    public string FactSelectValue
    {
        get
        {
            return ddl_fact.SelectedValue.ToString();
        }
        set
        {
            ddl_fact.SelectedValue = value;
        }
    }
    /// <summary>
    /// 部門
    /// value:dept_id
    /// </summary>
    public string DptSelectValue
    {
        get
        {
            return ddl_dpt.SelectedValue.ToString();
        }
        set
        {
            ddl_dpt.SelectedValue = value;
        }
    }
    /// <summary>
    /// 組別
    /// value:group_id
    /// </summary>
    public string GrpSelectValue
    {
        get
        {
            return ddl_group.SelectedValue.ToString();
        }
        set
        {
            ddl_group.SelectedValue = value;
        }
    }
    #endregion

    #region  ' select valueText '
    /// <summary>
    /// 取的公司別
     /// </summary>
    public string CompSelectValueTxt
    {
        get
        {
            return ddl_comp.SelectedItem.Text.Trim();
        }
        set
        {
            ddl_comp.SelectedItem.Text = value;
        }
    }
    /// <summary>
    /// 廠別
     /// </summary>
    public string FactSelectValueTxt
    {
        get
        {
            return ddl_fact.SelectedItem.Text.Trim();
        }
        set
        {
            ddl_fact.SelectedItem.Text = value;
        }
    }
    /// <summary>
    /// 部門
     /// </summary>
    public string DptSelectValueTxt
    {
        get
        {
            return ddl_dpt.SelectedItem.Text.Trim();
        }
        set
        {
            ddl_dpt.SelectedItem.Text = value;
        }
    }
    /// <summary>
    /// 組別
     /// </summary>
    public string GrpSelectValueTxt
    {
        get
        {
            return ddl_group.SelectedItem.Text.Trim();
        }
        set
        {
            ddl_group.SelectedItem.Text = value;
        }
    }
    #endregion

    #region ' ddl visible status '
     public bool ddlgroup_visible
    {
        get
        {
            return _ddlgroupVis;
        }
        set
        {
            _ddlgroupVis = value;

        }
    }

     public bool ddlDpt_visible
    {
        get
        {
            return _ddlDptVis;
        }
        set
        {
            _ddlDptVis = value;

        }
    }

     public bool ddlfact_visible
    {
        get
        {
            return _ddlfactVis;
        }
        set
        {
            _ddlfactVis = value;

        }
    }
    #endregion

    #endregion

    #region 'PageLoad'
     protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
             SetComp();
             chkVisforDDL();
             lbl_comp.Text = Resources.Strings.comp +"：";
             lbl_fact.Text = Resources.Strings.fact + "：";
             lbl_dept.Text = Resources.Strings.dept + "：";
             lbl_group.Text = Resources.Strings.group + "：";

             if (Request.QueryString["comp_no"] != "" && Request.QueryString["comp_no"] != null)
             {
                 ddl_comp.Items.FindByValue(Request.QueryString["comp_no"]).Selected = true;
                 ddl_SelectedIndexChanged((object)ddl_comp, e);
             }
        }
    }
    #endregion

    /// <summary>
    /// 取得公司資料
    /// </summary>
    public void SetComp()
    {
      
        PersonnelVN plvn = new PersonnelVN();
        //comp
        ddl_comp.Items.Clear();
        ddl_comp.DataSource = plvn.SimpleGetData("authcomp", null, "comp_cname");
        ddl_comp.DataTextField = "comp_cname";
        ddl_comp.DataValueField = "comp_no";
        ddl_comp.DataBind();
        ddl_comp.Items.Insert(0, new ListItem("---- Selected ----", "0"));

    }
    //ddl select change
    protected void ddl_SelectedIndexChanged(object sender, EventArgs e)
    {
        DropDownList _selectddl = sender as DropDownList;
        PersonnelVN plvn = new PersonnelVN();
        ArrayList al = new ArrayList();

        #region 'switch'
        switch (_selectddl.ID.Trim())
        {
            case "ddl_comp":
                ddl_fact.Items.Clear();
                if (ddl_comp.SelectedValue.Trim() != "0")
                {
                    if (!string.IsNullOrEmpty(ddl_comp.SelectedValue.Trim()))
                    {
                        al.Add("comp_no:" + ddl_comp.SelectedValue);
                    }
                    DataTable dt = plvn.SimpleGetData("authfact", al, "fact_no");
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        ListItem item = new ListItem();
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            item = new ListItem();
                            item.Text = Commfun.CheckDBNull(dt.Rows[i]["fact_no"]) + ":" + Commfun.CheckDBNull(dt.Rows[i]["fact_cname"]);
                            item.Value = Commfun.CheckDBNull(dt.Rows[i]["fact_no"]);
                            ddl_fact.Items.Add(item);
                        }
                    }
                    ddl_fact.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }
                else
                {
                    ddl_fact.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }

                ddl_dpt.Items.Clear();
                ddl_group.Items.Clear();
                ddl_dpt.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                ddl_group.Items.Insert(0, new ListItem("---- Selected ----", "0"));

                if (Request.QueryString["fact_no"] != "" && Request.QueryString["fact_no"] != null && !IsPostBack)
                {
                    ddl_fact.Items.FindByValue(Request.QueryString["fact_no"]).Selected = true;
                    ddl_SelectedIndexChanged((object)ddl_fact, e);
                }


                break;

            case "ddl_fact":
                ddl_dpt.Items.Clear();
                if (ddl_fact.SelectedValue.Trim() != "0")
                {
                    if (!string.IsNullOrEmpty(ddl_fact.SelectedValue.Trim()))
                    {
                        //al.Add("fact_no:" + ddl_fact.SelectedValue.Trim());
                        //al.Add("end_mk:N");

                        //ddl_dpt.DataSource = plvn.SimpleGetData("authdept", al, "");
                        ddl_dpt.DataSource = plvn.DataBySqlStr("SELECT id, dept_cname FROM   authdept  WHERE   (end_mk <> 'Y' OR end_mk is null OR end_mk='N')  AND  fact_no = '" + ddl_fact.SelectedValue + "' Order by dept_cname");
                        ddl_dpt.DataTextField = "dept_cname";
                        //ddl_dpt.DataValueField = "dept_no";       //  XKw 20090228    因為部門代號右重複的情況,所以改為dept_id來進行後續的關連
                        ddl_dpt.DataValueField = "id";
                        ddl_dpt.DataBind();
                    }
                    ddl_dpt.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }
                else
                {
                    ddl_dpt.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                    ddl_group.Items.Clear();
                    ddl_group.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }

                if (Request.QueryString["dept_no"] != "" && Request.QueryString["dept_no"] != null && !IsPostBack)
                {
                    ddl_dpt.Items.FindByValue(Request.QueryString["dept_no"]).Selected = true;
                    ddl_SelectedIndexChanged((object)ddl_dpt, e);
                }

                break;

            case "ddl_dpt":
                ddl_group.Items.Clear();
                if (ddl_dpt.SelectedValue.Trim() != "0")
                {

                    if (!string.IsNullOrEmpty(ddl_dpt.SelectedValue.Trim()))
                    {
                         al.Add("dept_no:" + ddl_dpt.SelectedValue.Trim());        //  XKw 20090228    因為部門代號有重複的情況,此處傳入的值已經為id
                        
                    }
                    ddl_group.DataSource = plvn.SimpleGetData("authgroup", al, "group_cname");
                    ddl_group.DataTextField = "group_cname";
                    ddl_group.DataValueField = "group_no";
                    ddl_group.DataBind();
                    ddl_group.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }
                else
                {
                    ddl_group.Items.Insert(0, new ListItem("---- Selected ----", "0"));
                }

                if (Request.QueryString["group_no"] != "" && Request.QueryString["group_no"] != null && !IsPostBack)
                {
                    ddl_group.Items.FindByValue(Request.QueryString["group_no"]).Selected = true;
                    ddl_SelectedIndexChanged((object)ddl_group, e);
                }

                break;
        }
        #endregion
    }
    /// <summary>
    /// 檢查ddl的顯示狀態為何
    /// </summary>
    protected void chkVisforDDL()
    {
        if (!_ddlgroupVis)
        {
            ddl_group.Visible = false;
            lbl_group.Visible = false;
        }

        if (!_ddlDptVis)
        {
            ddl_dpt.Visible = false;
            lbl_dept.Visible = false;
        }

        if (!_ddlfactVis)
        {
            ddl_fact.Visible = false;
            lbl_fact.Visible = false;
        }

    }
 
}

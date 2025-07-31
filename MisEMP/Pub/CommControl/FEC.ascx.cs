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
/// 要掛權限............
/// </summary>
public partial class Pub_CommControl_FEC : System.Web.UI.UserControl
{

    #region ' Variables '
    protected bool _ddlgroupVis = true;
    protected bool _ddlDptVis = true;
    protected bool _ddlfactVis = true;//預設顯示
    #endregion

    #region 'Public property value'

    #region ' select value '
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


    #region  ' ddlControl '
    /// <summary>
    /// 取的公司別
    /// </summary>
    public DropDownList Obj_ddl_comp
    {
        get
        {
            return ddl_comp;
        }
    }

    public DropDownList Obj_ddl_fact
    {
        get
        {
            return ddl_fact;
        }
    }
    public DropDownList Obj_ddl_dpt
    {
        get
        {
            return ddl_dpt;
        }
    }
    public DropDownList Obj_ddl_group
    {
        get
        {
            return ddl_group;
        }
    }


    #endregion

    #region  ' lblControl '
    public Label Obj_lbl_comp
    {
        get
        {
            return lbl_comp;
        }
    }
    public Label Obj_lbl_fact
    {
        get
        {
            return lbl_fact;
        }
    }
    public Label Obj_lbl_dept
    {
        get
        {
            return lbl_dept;
        }
    }

    public Label Obj_lbl_group
    {
        get
        {
            return lbl_group;
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

    string strUserID = "";
    protected void Page_Load(object sender, EventArgs e)
    {
        if (Session["UserID"] != null)
        {
            strUserID = Session["UserID"].ToString();
            strUserID = strUserID.Replace(".00", "").Replace(",00", "");
        }
        else
        {
            //RegisterClientScriptBlock("New", "<script language=javascript>alert('網頁已過期！(Time Out)');window.open('../../Default.aspx','_top');</script>");
            return;
        }

        if (!IsPostBack)
        {
            SetComp();
            chkVisforDDL();
            lbl_comp.Text = Resources.Strings.comp + "：";
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

    #region 'Event'
    //ddl select change
    protected void ddl_SelectedIndexChanged(object sender, EventArgs e)
    {
        DropDownList _selectddl = sender as DropDownList;
        PersonnelVN plvn = new PersonnelVN();
        string[] _getID = hvaeAuth();

        #region 'switch'
        switch (_selectddl.ID.Trim())
        {
            case "ddl_comp":
                ddl_fact.Items.Clear();
                if (ddl_comp.SelectedValue.Trim() != "0")
                {
                    if (!string.IsNullOrEmpty(_getID[1]))
                    {
                        // ddl_fact.DataSource = plvn.SimpleGetData("authfact", al, "");
                        DataTable dt = plvn.DataBySqlStr("SELECT * FROM   authfact  WHERE  fact_no IN (" + _getID[1] + ") AND  comp_no = '" + ddl_comp.SelectedValue + "' ORDER BY fact_no");
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
                    if (!string.IsNullOrEmpty(_getID[2]))
                    {
                        //ddl_dpt.DataSource = plvn.SimpleGetData("authdept", al, "");
                        ddl_dpt.DataSource = plvn.DataBySqlStr("SELECT id,dept_no+':'+dept_cname dept_cname FROM   authdept  WHERE   (end_mk <> 'Y' OR end_mk is null OR end_mk='N')  and id IN (" + _getID[2] + ") AND  fact_no = '" + ddl_fact.SelectedValue + "' ORDER BY dept_no");
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
                    if (!string.IsNullOrEmpty(_getID[3]))
                    {
                        // ddl_group.DataSource = plvn.SimpleGetData("authgroup", al, "");
                        ddl_group.DataSource = plvn.DataBySqlStr("SELECT group_no, group_cname FROM   authgroup  WHERE  group_no IN (" + _getID[3] + ") AND  dept_no = '" + ddl_dpt.SelectedValue + "' ORDER BY group_cname");
                        ddl_group.DataTextField = "group_cname";
                        ddl_group.DataValueField = "group_no";
                        ddl_group.DataBind();
                    }
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

    #endregion

    #region 'Function'
    /// <summary>
    /// 取得公司資料
    /// </summary>
    public void SetComp()
    {
        PersonnelVN plvn = new PersonnelVN();
        string[] _getID = hvaeAuth();

        if (!string.IsNullOrEmpty(_getID[0]))
        {
            //comp
            ddl_comp.Items.Clear();
            ddl_comp.DataSource = plvn.DataBySqlStr("SELECT * FROM   authcomp  WHERE  comp_no IN (" + _getID[0] + ")");
            ddl_comp.DataTextField = "comp_cname";
            ddl_comp.DataValueField = "comp_no";
            ddl_comp.DataBind();
        }
        ddl_comp.Items.Insert(0, new ListItem("---- Selected ----", "0"));

    }
    /// <summary>
    /// 檢查是否有資料在權限表
    ///  string[] _Ary 存放對應資料的 ID key
    /// _Ary[0] : comp 、_Ary[1] : fact 、_Ary[2] : dept 、_Ary[03] : group 
    /// </summary>
    /// <returns> 無權限時 ddl全鎖 </returns>
    protected string[] hvaeAuth()
    {
        string[] _Ary = new string[4];
        string _compid = "", _cmpidStr = "", _factid = "", _factidStr = "", _deptid = "", _deptidStr = "", _groupid = "", _groupidStr = "";

        if (Session["UserID"] != null)
        {
            strUserID = Session["UserID"].ToString();
            strUserID = strUserID.Replace(".00", "").Replace(",00", "");
        }
        else
        {
            //RegisterClientScriptBlock("New", "<script language=javascript>alert('網頁已過期！(Time Out)');window.open('../../Default.aspx','_top');</script>");
            return _Ary;
        }
        PersonnelVN plvn = new PersonnelVN();        
        DataTable _dt = plvn.Get_UserAllauthority(strUserID);   //get user all authority

        if (_dt.Rows.Count == 0) //  無權限直接鎖ddl
        {
            ddl_comp.Enabled = false;
            ddl_fact.Enabled = false;
            ddl_dpt.Enabled = false;
            ddl_group.Enabled = false;
        }
        else
        {

            foreach (DataRow dr in _dt.Select("", "comp_no"))
            {
                #region get comp
                if (_compid.CompareTo(dr["comp_no"].ToString().Trim()) != 0)
                {
                    if (string.IsNullOrEmpty(_cmpidStr))
                    {
                        _cmpidStr = "'" + dr["comp_no"].ToString().Trim() + "'";
                    }
                    else
                    {
                        _cmpidStr = _cmpidStr + "," + "'" + dr["comp_no"].ToString().Trim() + "'";
                    }

                    _compid = dr["comp_no"].ToString().Trim();
                }
                #endregion

                #region get fact
                if (_factid.CompareTo(dr["fact_no"].ToString().Trim()) != 0)
                {
                    if (string.IsNullOrEmpty(_factidStr))
                    {
                        _factidStr = "'" + dr["fact_no"].ToString().Trim() + "'";
                    }
                    else
                    {
                        _factidStr = _factidStr + "," + "'" + dr["fact_no"].ToString().Trim() + "'";
                    }

                    _factid = dr["fact_no"].ToString().Trim();
                }
                #endregion

                #region  get dept

                if (_deptid.CompareTo(dr["deptid"].ToString().Trim()) != 0)
                {
                    if (string.IsNullOrEmpty(_deptidStr))                              //   第一筆的時候
                    {
                        // _deptidStr = "'" + dr["dept_no"].ToString().Trim() + "'";
                        _deptidStr = dr["deptid"].ToString().Trim();        //  XKw 20090228    因為部門代號右重複的情況,所以改為dept_id來進行後續的關連, id為numeric
                    }
                    else
                    {
                        //_deptidStr = _deptidStr + "," + "'" + dr["dept_no"].ToString().Trim() + "'"; 
                        _deptidStr += "," + dr["deptid"].ToString().Trim();         //  id為numeric 所以不需要用'號
                    }

                    _deptid = dr["deptid"].ToString().Trim();
                }
                #endregion

                #region get group
                //get group
                if (_groupid.CompareTo(dr["group_no"].ToString().Trim()) != 0)
                {
                    if (string.IsNullOrEmpty(_groupidStr))
                    {
                        _groupidStr = "'" + dr["group_no"].ToString().Trim() + "'";
                    }
                    else
                    {
                        _groupidStr = _groupidStr + "," + "'" + dr["group_no"].ToString().Trim() + "'";
                    }

                    _groupid = dr["group_no"].ToString().Trim();
                }
                #endregion
            }

        }
        _Ary[0] = _cmpidStr;
        _Ary[1] = _factidStr;
        _Ary[2] = _deptidStr;
        _Ary[3] = _groupidStr;

        return _Ary;
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
    #endregion
}

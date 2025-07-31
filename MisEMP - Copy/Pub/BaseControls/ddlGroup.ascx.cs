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

public partial class Pub_BaseControls_ddlGroup : System.Web.UI.UserControl
{

    #region '共用基本屬性'
    public string SelectedValue
    {
        get
        {
            return ddlGroup.SelectedValue.ToString();
        }
        set
        {
            ddlGroup.SelectedValue = value;
        }
    }
    public ListItemCollection Items
    {
        get
        {
            return ddlGroup.Items;
        }

    }
    public bool enable
    {
        get
        {
            return ddlGroup.Enabled;
        }
        set
        {
            ddlGroup.Enabled = value;
        }
    }
    public bool visible
    {
        get
        {
            return ddlGroup.Visible;
        }
        set
        {
            ddlGroup.Visible = value;
        }
    }
    /// <summary>
    /// ture 為selectedindexChangse 啟用
    /// </summary>
    public bool AutoPostBk
    {
        get
        {
            return ddlGroup.AutoPostBack;
        }
        set
        {
            ddlGroup.AutoPostBack = value;
        }
    }
    #endregion

    protected void Page_Load(object sender, EventArgs e)
    {

    }

    /// <summary>
    /// 取得對應的組別
    /// </summary>
    /// <param name="comp_id">群別id, 如要取全部-不分部門, 帶空字串</param>
    public void GetGroup(string dept_id)
    {
        ddlGroup.Items.Clear();
        PersonnelVN plvn = new PersonnelVN();
        ArrayList al = new ArrayList();
        if (dept_id.Trim() != "0")
        {
            if (!string.IsNullOrEmpty(dept_id))
            {
                al.Add("dept_id:" + dept_id.Trim());
            }
            
            //fact 
            ddlGroup.DataSource = plvn.SimpleGetData("authgroup", al, "");
            ddlGroup.DataTextField = "group_cname";
            ddlGroup.DataValueField = "ID";
            ddlGroup.DataBind();
            ddlGroup.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }
        else
        {
            ddlGroup.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }

    }
}

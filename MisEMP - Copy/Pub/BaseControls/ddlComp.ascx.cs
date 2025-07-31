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

 
public partial class Pub_BaseControls_ddlComp : System.Web.UI.UserControl
{
    #region 'PageLoad'
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    #endregion

    #region '共用基本屬性'
    public string SelectedValue
    {
        get
        {
            return ddlComp.SelectedValue.ToString();
        }
        set
        {
            ddlComp.SelectedValue = value;
        }
    }
    public ListItemCollection Items
    {
        get
        {
            return ddlComp.Items;
        }

    }
    public bool enable
    {
        get
        {
            return ddlComp.Enabled;
        }
        set
        {
            ddlComp.Enabled = value;
        }
    }
    public bool visible
    {
        get
        {
            return ddlComp.Visible;
        }
        set
        {
            ddlComp.Visible = value;
        }
    }
    /// <summary>
    /// ture 為selectedindexChangse 啟用
    /// </summary>
    public bool AutoPostBk
    {
        get
        {
            return ddlComp.AutoPostBack;
        }
        set
        {
            ddlComp.AutoPostBack = value;
        }
    }

 
    #endregion

    /// <summary>
    /// 取得公司
    /// </summary>
    public void GetComp()
    {
        PersonnelVN plvn = new PersonnelVN();
        //comp
        ddlComp.DataSource = plvn.SimpleGetData("authcomp", null, "");
        ddlComp.DataTextField = "comp_cname";
        ddlComp.DataValueField = "ID";
        ddlComp.DataBind();
    }
   
}




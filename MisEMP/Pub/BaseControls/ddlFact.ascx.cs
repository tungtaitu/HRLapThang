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
 

 
public partial class Pub_BaseControls_ddlFact : System.Web.UI.UserControl
{
    
    #region '共用基本屬性'
    public string SelectedValue
    {
        get
        {
            return ddlFact.SelectedValue.ToString();
        }
        set
        {
            ddlFact.SelectedValue = value;
        }
    }
    public ListItemCollection Items
    {
        get
        {
            return ddlFact.Items;
        }

    }
    public bool enable
    {
        get
        {
            return ddlFact.Enabled;
        }
        set
        {
            ddlFact.Enabled = value;
        }
    }
    public bool visible
    {
        get
        {
            return ddlFact.Visible;
        }
        set
        {
            ddlFact.Visible = value;
        }
    }
    /// <summary>
    /// ture 為selectedindexChangse 啟用
    /// </summary>
    public bool AutoPostBk
    {
        get
        {
            return ddlFact.AutoPostBack;
        }
        set
        {
            ddlFact.AutoPostBack = value;
        }
    }
    #endregion

    #region 'PageLoad'
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    #endregion

    /// <summary>
    /// 取得對應的廠別
    /// </summary>
    /// <param name="comp_id">公司別id, 如要取全部廠不分公司帶空字串</param>
    public void GetFact(string comp_id)
    {
        ddlFact.Items.Clear();
        PersonnelVN plvn = new PersonnelVN();
        ArrayList al = new ArrayList();
        if (comp_id.Trim() != "0")
        {
            if (!string.IsNullOrEmpty(comp_id))
            {
                al.Add("comp_id:" + comp_id.Trim());
            }

            //fact 
            ddlFact.DataSource = plvn.SimpleGetData("authfact", al, "");
            ddlFact.DataTextField = "fact_cname";
            ddlFact.DataValueField = "ID";
            ddlFact.DataBind();
            ddlFact.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }
        else
        {
            ddlFact.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }

    }
    
}
 
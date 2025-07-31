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

public partial class Pub_BaseControls_ddlDept : System.Web.UI.UserControl
{

    #region '共用基本屬性'
    public string SelectedValue
    {
        get
        {
            return ddlDept.SelectedValue.ToString();
        }
        set
        {
            ddlDept.SelectedValue = value;
        }
    }
    public ListItemCollection Items
    {
        get
        {
            return ddlDept.Items;
        }

    }
    public bool enable
    {
        get
        {
            return ddlDept.Enabled;
        }
        set
        {
            ddlDept.Enabled = value;
        }
    }
    public bool visible
    {
        get
        {
            return ddlDept.Visible;
        }
        set
        {
            ddlDept.Visible = value;
        }
    }
    /// <summary>
    /// ture 為selectedindexChangse 啟用
    /// </summary>
    public bool AutoPostBk
    {
        get
        {
            return ddlDept.AutoPostBack;
        }
        set
        {
            ddlDept.AutoPostBack = value;
        }
    }
    #endregion

    #region 'PageLoad'
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    #endregion

    /// <summary>
    /// 取得對應的部門
    /// </summary>
    /// <param name="comp_id">部門別id, 如要取全部-部門不分廠別, 帶空字串</param>
    public void GetDept(string fact_id)
    {
        ddlDept.Items.Clear();
        PersonnelVN plvn = new PersonnelVN();
        ArrayList al = new ArrayList();
        if (fact_id.Trim() != "0")
        {
            if (!string.IsNullOrEmpty(fact_id))
            {
                al.Add("fact_id:" + fact_id.Trim());
            }

            //fact 
            ddlDept.DataSource = plvn.SimpleGetData("authdept", al, "");
            ddlDept.DataTextField = "dept_cname";
            ddlDept.DataValueField = "ID";
            ddlDept.DataBind();
            ddlDept.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }
        else
        {
            ddlDept.Items.Insert(0, new ListItem("---- 請選擇 ----", "0"));
        }

    }
    


}

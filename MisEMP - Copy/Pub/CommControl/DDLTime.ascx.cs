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

public partial class Pub_CommControl_DDLTime : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {


    }
    /// <summary>
    /// 小時
    /// </summary>
    public string H_SelectedValue
    {
        get
        {
            return ddlH.SelectedValue.ToString();
        }
        set
        {
            ddlH.SelectedValue = value;
        }
    }
    /// <summary>
    /// 分鐘
    /// </summary>
    public string M_SelectedValue
    {
        get
        {
            return ddlM.SelectedValue.ToString();
        }
        set
        {
            ddlM.SelectedValue = value;
        }
    }

    public bool Enabled
    {
        set
        {
            ddlH.Enabled = value;
            ddlM.Enabled = value;
        }
    }
}

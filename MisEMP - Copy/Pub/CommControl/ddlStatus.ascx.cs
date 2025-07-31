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

public partial class Pub_CommControl_ddlStatus : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    public string SelectedValue
    {
        get
        {
            return ddlStatus.SelectedValue.ToString();
        }
        set
        {
            ddlStatus.SelectedValue = value;
        }
    }
    
}

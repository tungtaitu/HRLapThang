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

public partial class Pub_CommControl_ddlMissCardType : System.Web.UI.UserControl
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }
    public void GetMissType()
    {
        PersonnelVN PVN = new PersonnelVN();
        ArrayList AL = new ArrayList();
        AL.Add("Customname:MissCardType");
        AL.Add("endmk:N");
        DataTable dt = PVN.SimpleGetData("custonInfo", AL, "");
        DropDownList1.DataSource = dt;
        DropDownList1.DataTextField = "CustomDesc";
        DropDownList1.DataValueField = "CustomValue";
        DropDownList1.DataBind();
    }

    public string SelectedValue
    {
        get
        {
            return DropDownList1.SelectedValue.ToString();
        }
        set
        {
            DropDownList1.SelectedValue = value;
        }
    }
    public bool Enabled
    {
        set
        {
            DropDownList1.Enabled = value;
        }
    }
}

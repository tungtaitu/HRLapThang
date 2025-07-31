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
using PccCommonForC;

public partial class Pub_Module_ShowSignPicture : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

        //WebPUB
        db_OverApplyGroup1 Over = new db_OverApplyGroup1(ConfigurationManager.AppSettings["ConnectionType"], ConfigurationManager.AppSettings["ConnectionServer"], ConfigurationManager.AppSettings["ConnectionDB"], ConfigurationManager.AppSettings["ConnectionUser"], ConfigurationManager.AppSettings["ConnectionPwd"]);
        DataSet ds = Over.GetImageFile(Request.QueryString["SignId"]); // signId la user_id
        DataTable ImageTable = new DataTable();
        ImageTable = ds.Tables["ImageFile"];
        byte[] byteReturn = { };
        if (ImageTable.Rows.Count > 0)
        {
            byteReturn = (byte[])ImageTable.Rows[0]["sign_pic"];
            Response.BinaryWrite(byteReturn);

        }
    }

    #region Web Form 設計工具產生的程式碼
    override protected void OnInit(EventArgs e)
    {
        //
        // CODEGEN: 此為 ASP.NET Web Form 設計工具所需的呼叫。
        //
        InitializeComponent();
        base.OnInit(e);
    }

    /// <summary>
    /// 此為設計工具支援所必須的方法 - 請勿使用程式碼編輯器修改
    /// 這個方法的內容。
    /// </summary>
    private void InitializeComponent()
    {
        this.Load += new System.EventHandler(this.Page_Load);
    }
    #endregion
}

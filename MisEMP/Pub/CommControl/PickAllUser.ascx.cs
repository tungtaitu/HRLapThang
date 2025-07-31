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
/// 挑選所有在vie_employee的人員，不過濾
/// </summary>
public partial class Pub_CommControl_PickAllUser : System.Web.UI.UserControl
{
    #region Public Property
    /// <summary>
    /// 在TxtBox顯示的userNm
    /// 只要將　ShowTextBoxUserNm = ChoiceUserNm
    /// txtBox即會顯示挑選的使用者名字。
    /// </summary>
    public string ShowTextBoxUserNm
    {
        get
        {
            return txt_UserInfo.Text;
        }
        set
        {
            txt_UserInfo.Text = value;
        }
    }
    /// <summary>
    ///  取得挑選的UserId
    /// </summary>
    public string ChoiceUserID
    {
        get
        {
            return txtReturnID.Value;
        }
        set
        {
            txtReturnID.Value = value;
        }
    }

    /// <summary>
    /// 取得挑選的UserName
    /// </summary>
    public string ChoiceUserNm
    {
        get
        {
            return hidReturnNm.Value;
        }
        set
        {
            hidReturnNm.Value = value;
        }
    }

    public Unit txtBoxWidth
    {
        get
        {
            return txt_UserInfo.Width;
        }
        set
        {
            txt_UserInfo.Width = value;
        }

    }

    public Unit txtBoxHeight
    {
        get
        {
            return txt_UserInfo.Height;
        }
        set
        {
            txt_UserInfo.Height = value;
        }


    }
    #endregion
    protected void Page_Load(object sender, EventArgs e)
    {

    }
}

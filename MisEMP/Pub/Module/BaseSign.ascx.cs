using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.ComponentModel;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Drawing;


public partial class Pub_Module_BaseSign : System.Web.UI.UserControl
{
    protected Table tblSign;
    private string ShowPicturePath = "ShowSignPicture.aspx?SignId=";
    private string ShowImagePath = "ShowSignImage.aspx?SignId=";

    #region 設定私有變數

    private string m_Users = string.Empty;
    private string m_Roles = string.Empty;
    private string m_SignDates = string.Empty;
    private string m_Auths = string.Empty;

    #endregion

    #region 設定共用屬性

    public string Users
    {
        get
        {
            return m_Users;
        }
        set
        {
            m_Users = value;
        }
    }

    public string Roles
    {
        get
        {
            return m_Roles;
        }
        set
        {
            m_Roles = value;
        }
    }

    public string SignDates
    {
        get
        {
            return m_SignDates;
        }
        set
        {
            m_SignDates = value;
        }
    }

    public string Authticates
    {
        get
        {
            return m_Auths;
        }
        set
        {
            m_Auths = value;
        }
    }

    #endregion


    private void Page_Load(object sender, System.EventArgs e)
    {
        if (m_Roles.Equals(string.Empty)) return;

        tblSign = new Table();

        if (m_Roles.Equals(string.Empty)) return;

        string[] users = m_Users.Split(',');
        string[] roles = m_Roles.Split(',');
        string[] signDates = m_SignDates.Split(',');
        string[] auths = m_Auths.Split(',');

        int picCount = 8;
        try
        {
            picCount = int.Parse(this.Attributes["PicCount"]);
        }
        catch
        {
            picCount = 8;
        }

        int maxPercent = (int)Math.Floor((Double)(100 / picCount));


        TableRow row = new TableRow();
        TableCell cell;


        Table tblC = new Table();
        TableRow rowC;
        TableCell cellC;
        tblC.BorderColor = Color.Black;
        tblC.BorderWidth = Unit.Pixel(1);






        int spacePercent, signPercent;
        signPercent = countSignPercent(roles.Length, picCount, out spacePercent);
        int widPercent = maxPercent * roles.Length;

        string user_id = string.Empty;
        string signDate = string.Empty;
        string auth = string.Empty;




        if (spacePercent != 100)
        {
            int countSign = 0;
            countSign = roles.Length;

            bool flagExcess = false;
            float iExcess = 0;
            if (countSign < picCount) // so nguoi ky nho hon so cot
            {
                iExcess = picCount - countSign;
                flagExcess = true;
            }
            else if (countSign > picCount)
            {
                iExcess = picCount * 2 - countSign;
                flagExcess = true;
            }
            if (flagExcess == true)
            {
                row = new TableRow();
                for (int i = 0; i < iExcess; i++)
                {

                    Table tblE = new Table();
                    rowC = new TableRow();
                    cellC = new TableCell();
                    cellC.Text = "<br>";
                    rowC.Cells.Add(cellC);
                    tblE.Rows.Add(rowC);
                    tblE.Width = Unit.Percentage(100);

                    cell = new TableCell();
                    cell.Controls.Add(tblE);
                    //cell.Width=Unit.Percentage(maxPercent);
                    cell.Width = Unit.Pixel(100); // phi Add 
                    row.Cells.Add(cell);

                }
                tblParent.Rows.Add(row);
            }





            int icol = int.Parse(iExcess.ToString());
            for (int i = roles.Length - 1; i > -1; i--)
            {
                icol++;


                //先判斷是否有相同的角色及使用者數目
                if (i > users.Length - 1)
                {
                    user_id = "0";
                }
                else
                {
                    if (!users[i].Equals(string.Empty))
                        user_id = users[i];
                    else
                        user_id = "0";
                }

                //判斷是否有其簽核日期 
                if (i > signDates.Length - 1)
                    signDate = "<br>";
                else
                {
                    if (!signDates[i].Equals(string.Empty))
                        signDate = signDates[i];
                    else
                        signDate = "<br>";
                }

                //把是否授權的部份指定
                if (i > auths.Length - 1)
                    auth = "N";
                else
                {
                    if (!auths[i].Equals(string.Empty))
                        auth = auths[i];
                    else
                        auth = "N";
                }


                string MasterStyle = string.Empty;

                tblC = new Table();

                rowC = new TableRow();
                addRowCss(rowC); //phi Add
                cellC = new TableCell();
                cellC.Text = "<b>" + roles[i] + "</b>";
                cellC.Wrap = false;
                addCss(cellC);// Phi Add
              //  cellC.CssClass = "TableRowStyle"; // Phi edit
                cellC.Height = Unit.Pixel(20);
               // cellC.Width = Unit.Pixel(102);
                rowC.Cells.Add(cellC);
                tblC.Rows.Add(rowC);

                rowC = new TableRow();
                addRowCss(rowC);// Phi Add
                cellC = new TableCell();
                System.Web.UI.WebControls.Image myImage = new System.Web.UI.WebControls.Image();
                myImage.Width = Unit.Pixel(100);
                myImage.Height = Unit.Pixel(100);
                
                // end edit
                if (user_id != "0" && CheckSignImageFile(user_id) == true)
                {
                    if (auth.Equals("N"))
                        myImage.ImageUrl = ShowPicturePath + user_id;
                    else
                        myImage.ImageUrl = ShowImagePath + user_id;
                }
                else
                {
                    myImage.ImageUrl = ShowImagePath + user_id;
                }


                if (user_id != "0")
                {
                    cellC.Controls.Add(myImage);
                }
                else
                {
                    cellC.Text = "<b>";
                }
                rowC.Cells.Add(cellC);
                addCss(cellC);// phi edit
               // cellC.CssClass = "TableRowStyle";
                cellC.Height = Unit.Pixel(102);
                cellC.Width = Unit.Pixel(102);
                tblC.Rows.Add(rowC);

                rowC = new TableRow();
                addRowCss(rowC);
                cellC = new TableCell();
                cellC.Text = signDates[i];
                cellC.Wrap = false;
                addCss(cellC);// phi edit
              //  cellC.CssClass = "TableRowStyle";
                cellC.Height = Unit.Pixel(20);
                rowC.Cells.Add(cellC);
                tblC.Rows.Add(rowC);

                tblC.CellPadding = 0;
                tblC.CellSpacing = 0;
                tblC.Width = Unit.Percentage(100);
                tblC.Height = Unit.Percentage(100);

                if (flagExcess != true || icol - 1 == picCount) //
                {
                    row = new TableRow();
                    row.Height = Unit.Pixel(5);
                    tblParent.Rows.Add(row);
                    flagExcess = true;
                    row = new TableRow();
                }

                cell = new TableCell();
                cell.Controls.Add(tblC);
                //cell.Width=Unit.Percentage(maxPercent);
                row.Cells.Add(cell);
                tblParent.Rows.Add(row);
                tblParent.BorderWidth = Unit.Pixel(0);
                tblParent.CellPadding = 0;
                tblParent.CellSpacing = 0;

                //tblParent.Width=Unit.Percentage(100);


            }
        }

        m_Roles = string.Empty;
    }



    private int countSignPercent(int roleLength, int picCount, out int spacePercent)
    {
        if (roleLength == 0)
        {
            spacePercent = 100;
            return 0;
        }
        double d1 = 100 / roleLength;
        int i = (int)Math.Ceiling(d1);

        //計算最大的Percent是多少？ 20080724 by Lemor
        int maxPercent = (int)Math.Floor((Double)(100 / picCount));

        if (i > maxPercent)
        {
            i = maxPercent;
            spacePercent = 100 - (i * roleLength);
        }
        else
        {
            spacePercent = 0;
        }

        return i;
    }

    private bool CheckSignImageFile(string user_id)
    {
        //GetImageFile
        db_OverApplyGroup1 Over = new db_OverApplyGroup1(ConfigurationManager.AppSettings["ConnectionType"], ConfigurationManager.AppSettings["ConnectionServer"], ConfigurationManager.AppSettings["ConnectionDB"], ConfigurationManager.AppSettings["ConnectionUser"], ConfigurationManager.AppSettings["ConnectionPwd"]);
        DataSet ds = Over.GetImageFile(user_id);
        DataTable ImageTable = new DataTable();
        ImageTable = ds.Tables["ImageFile"];
        bool flag = false;
        if (ImageTable.Rows.Count > 0)
            flag = true;
        return flag;
    }

    private void addRowCss(TableRow row)
    {
        row.HorizontalAlign = HorizontalAlign.Center;
        row.VerticalAlign = VerticalAlign.Middle;
    }

    private void addCss(TableCell cell)
    {        
        cell.HorizontalAlign = HorizontalAlign.Center;
        cell.VerticalAlign = VerticalAlign.Middle;
        cell.Style["border"] = "black 1px solid";
    }

    #region 產生一格簽名表格的程式

    private Table genSignTable(string user_id, string roleName, string signDate, string auth)
    {
        Table myTable = new Table();
        myTable.Attributes.Add("border", "2");

        //設定簽名檔
        System.Web.UI.WebControls.Image myImage = new System.Web.UI.WebControls.Image();
        if (user_id != "0")
        {
            if (auth.Equals("N"))
                myImage.ImageUrl = ShowPicturePath + user_id;
            else
                myImage.ImageUrl = ShowImagePath + user_id;
        }

        //加入角色

        PLSignWeb2.Pub.Module.PccRow myRow = new PLSignWeb2.Pub.Module.PccRow();
        myRow.SetDefaultCellData(string.Empty, HorizontalAlign.Center, VerticalAlign.Middle, 0);
        myRow.AddTextCell(getFormatRoleNm(roleName), 3);
        if (user_id != "0")
        {
            myRow.AddControl(myImage, 97);
        }
        else
        {
            myRow.AddTextCell("<br>", 97);
        }

        myTable.Rows.Add(myRow.Row);

        myTable.Rows.Add(getSignDateRow(signDate).Row);

        return myTable;
    }

    private PLSignWeb2.Pub.Module.PccRow getSignDateRow(string signDate)
    {
        PLSignWeb2.Pub.Module.PccRow myRow = new PLSignWeb2.Pub.Module.PccRow();
        myRow.SetDefaultCellData(string.Empty, HorizontalAlign.Center, VerticalAlign.Middle, 2);
        myRow.AddTextCell(signDate, 100);
        return myRow;
    }

    private string getFormatRoleNm(string roleName)
    {
        string strReturn = string.Empty;
        if (roleName.Equals(string.Empty))
        {
            strReturn = "<br><br><br><br><br><br><br><br>";
        }
        strReturn = "<b>";
        switch (roleName.Length)
        {
            case 0:
                strReturn += "<br><br><br><br><br><br><br>";
                break;
            case 1:
                strReturn += roleName + "<br><br><br><br><br><br>";
                break;
            case 2:
                strReturn += roleName.Substring(0, 1) + "<br><br>" + roleName.Substring(1, 1) + "<br><br>";
                break;
            case 3:
                strReturn += roleName + "<br><br>";
                break;
            case 4:
                strReturn += roleName;
                break;
            default:
                strReturn += roleName.Substring(0, 4);
                break;
        }

        strReturn += "</b>";

        return strReturn;
    }

    #endregion

    #region Web Form 設計工具產生的程式碼
    override protected void OnInit(EventArgs e)
    {
        //
        // CODEGEN: 此為 ASP.NET Web Form 設計工具所需的呼叫。
        //

        //計算這個網頁所在的層次是那裡
        int i, j = 0;
        string strPageLayer = "";
        string LocalPath = PccCommonForC.PccToolFunc.Upper(Server.MapPath("."));

        j = LocalPath.IndexOf(PccCommonForC.PccToolFunc.Upper(Application["EDPNET"].ToString()));

        try
        {
            for (i = 1; i < LocalPath.Substring(j).Split('\\').Length; i++)
            {
                strPageLayer += "../";
            }
            Session["PageLayer"] = strPageLayer;
        }
        catch
        {
            Session["PageLayer"] = "";
        }

        InitializeComponent();
        base.OnInit(e);
    }

    /// <summary>
    ///		此為設計工具支援所必須的方法 - 請勿使用程式碼編輯器修改
    ///		這個方法的內容。
    /// </summary>
    /// 
    private void InitializeComponent()
    {
        this.Load += new System.EventHandler(this.Page_Load);

    }
    #endregion
}
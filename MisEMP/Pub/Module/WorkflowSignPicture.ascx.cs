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
using System.Text;

public partial class Pub_Module_WorkflowSignPicture : System.Web.UI.UserControl
{
    private string m_Vou_No = string.Empty;  

    public string Vou_No
    {
        get
        {
            return m_Vou_No;
        }
        set
        {
            m_Vou_No = value;
        }
    }
   

    protected void Page_Load(object sender, EventArgs e)
    {
            string strUser = "";
            string strRoles = "";
            string strTime = "";
            string strAuthticates = "";
            string start_mk = "Y";
            bool first = true;
            DataSet dsAll = null;
            DataTable dtAll = null;
            DataSet ds = null;
            DataTable dt = null;
                      
            PccMsg myMsg = new PccMsg();
           
            //lay nguoi ki cua TTHC
            db_OverApplyGroup1 db_overApp = new db_OverApplyGroup1(ConfigurationManager.AppSettings["PLVNConnectionType"], ConfigurationManager.AppSettings["PLVNConnectionServer"], ConfigurationManager.AppSettings["PLVNConnectionDB"], ConfigurationManager.AppSettings["PLVNConnectionUser"], ConfigurationManager.AppSettings["PLVNConnectionPwd"]);
            
            ds = db_overApp.Get_Sign_Data(start_mk, m_Vou_No);
            dt = ds.Tables["SignData"];

            DataRow[] foundRowsSort;
            foundRowsSort = dt.Select("1=1", "SIGN_DATE");

                    if (dt.Rows.Count > 0)
                    {
                        foreach (DataRow myMasterRow in foundRowsSort)
                            {
                                strUser = strUser + CheckDBNull(myMasterRow["user_id"]).Trim() + ",";//人員
                                if (first)
                                {
                                    strRoles += "申請者,";//申請者
                                }
                                else
                                {
                                    if (CheckDBNull(myMasterRow["role_nm"]).Trim() != "")
                                    {
                                        strRoles += CheckDBNull(myMasterRow["role_nm"]).Trim() + ",";//職稱
                                    }
                                    else
                                        strRoles += " " + ",";//職稱
                                }
                                string strSignDate = CheckDBNull(myMasterRow["sign_date"]).Trim();
                                if (strSignDate.Length == 14)
                                    strSignDate = strSignDate.Substring(4, 2) + "/" + strSignDate.Substring(6, 2) + " " + strSignDate.Substring(8, 2) + ":" + strSignDate.Substring(10, 2);

                                //代理人的話，日期後加(代) 20060911
                                if (CheckDBNull(myMasterRow["sign_mk"]).Trim() == "A")//代理
                                    strTime = strTime + strSignDate + "&nbsp;&nbsp;" + "<font color=#ff3333 style='FONT-WEIGHT: bold'>(代)</font>" + ",";//簽核時間+(代)
                                else
                                    strTime = strTime + strSignDate + ",";//簽核時間

                                //strAuthticates = strAuthticates + "N,";//受權
                                //判斷 代理(A -->目前沒有)、受權(B)、正常(N) 20060909
                                if (CheckDBNull(myMasterRow["sign_mk"]).Trim() == "B")//受權
                                {
                                    strAuthticates = strAuthticates + "Y,";//受權
                                }
                                else
                                {
                                    strAuthticates = strAuthticates + "N,";//無受權
                                }
                                first = false;
                            }
                      }
                

            if (strUser != "")
                  strUser = strUser.Substring(0, strUser.Length - 1); //"176.00,301.00,0,0"
            if (strRoles != "")
                  strRoles = strRoles.Substring(0, strRoles.Length - 1); //"組員,核准主管,專案主管,專案主管"
            if (strTime != "")
                  strTime = strTime.Substring(0, strTime.Length - 1);
            if (strAuthticates != "")
                  strAuthticates = strAuthticates.Substring(0, strAuthticates.Length - 1);//"N,N,N,N"

                  BaseSign1.Users = strUser;
                  BaseSign1.Roles = strRoles;
                  BaseSign1.SignDates = strTime; //"02/29 16:07,02/29 16:51,,"
                  BaseSign1.Authticates = strAuthticates;

        }


    private void MasterTable(string user_desc, string role_nm, string note)
    {
        if (note.Trim() != "")
        {
            //加入主要的表格資訊 ================================================================

            PLSignWeb2.Pub.Module.PccRow myRow = new PLSignWeb2.Pub.Module.PccRow("DGridTD", HorizontalAlign.Center, VerticalAlign.Middle, 15);
            myRow.SetDefaultCellData(string.Empty, HorizontalAlign.Left, VerticalAlign.Middle, 0);

            myRow.AddTextCell(user_desc, 10);
            myRow.AddTextCell("(" + role_nm + ")：", 15);
            myRow.AddTextCell(note, 75);
            mTable.Rows.Add(myRow.Row);
        }
    }


        private string CheckQueryString(string strName)
        {
            if (Request.QueryString[strName] == null)
                return "";
            else
                return Request.QueryString[strName].ToString();
        }

        private string CheckDBNull(object oFieldData)
        {
            if (Convert.IsDBNull(oFieldData))
                return "";
            else
                return oFieldData.ToString().Trim();
        }

        #region Web Form Designer generated code
        override protected void OnInit(EventArgs e)
        {
            //
            // CODEGEN: This call is required by the ASP.NET Web Form Designer.
            //
            InitializeComponent();
            base.OnInit(e);
        }

        /// <summary>
        ///		Required method for Designer support - do not modify
        ///		the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Load += new System.EventHandler(this.Page_Load);

        }
        #endregion
    }

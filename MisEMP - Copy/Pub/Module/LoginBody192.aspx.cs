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
using Pcc.Utils.CollectionUtils;
using System.Threading;
using System.Globalization;
using PccBsLayerForC;
using PccCommonForC;
using System.Xml;
using System.Web.Configuration; 

public partial class Pub_Module_LoginBody192 : BasePage
{
    PersonnelVN PVN = new PersonnelVN();
    OracleFn ORA = new OracleFn();
    WebPub OWebPub = new WebPub();
    string strUserID = "";
    private const string MENUXML = "../XmlDoc/ApMenu.xml";
    private const string MENUXML_VN = "../XmlDoc/ApMenu_VN.xml";
    protected void Page_Load(object sender, EventArgs e)
    {
        /*
        if (WebConfigurationManager.AppSettings["Area"].ToString().Trim() != "PFQ")
        {
            Div11.Visible = false;
        }
        */

        Div1.Visible = false;
        Div2.Visible = false;
        Div3.Visible = false;
        Div4.Visible = false;
        Div5.Visible = false;
        Div6.Visible = false;
        Div7.Visible = false;
        
        /*
        string sArea=WebConfigurationManager.AppSettings["Area"].ToString().Trim();
        if (sArea == "PSV" || sArea == "PCV" || sArea == "PFQ" || sArea == "YBH")
        {
            Div10.Visible = true;
        }
        else {
            Div10.Visible = false;
        }
        */

        if (Session["UserID"] != null)
        {
            strUserID = Session["UserID"].ToString();
            strUserID = strUserID.Replace(".00", "").Replace(",00", "");
        }
        else
        {
            RegisterClientScriptBlock("New", "<script language=javascript>alert('網頁已過期！(Time Out)');window.open('../../Default.aspx','_top');</script>");
            return;
        }


        if (!IsPostBack)
        {
            //ShowCheckScreen();           
        }
    }
    /// <summary>
    /// 預設檢查的資料：sid、授權、代理、目前送審的單子
    /// </summary>
    
    private void ShowCheckScreen()
    {
        DataTable dt = GetCountData();
        if(dt.Rows.Count>0)
        {
            int Labsent = 0, OverTime = 0, MissCard = 0, OverTime_Leave = 0, AppOverTime = 0, LossCard=0, LevOut=0;

            Labsent = Convert.ToInt32(dt.Rows[0]["Labsent"]);
            OverTime = Convert.ToInt32(dt.Rows[0]["OverTime"]);
            MissCard = Convert.ToInt32(dt.Rows[0]["MissCard"]);
            OverTime_Leave = Convert.ToInt32(dt.Rows[0]["OverTime_Leave"]);
            try
            {
                LevOut = Convert.ToInt32(dt.Rows[0]["Lev_out"]);
            }
            catch
            {
            }
            try
            {
                AppOverTime = Convert.ToInt32(dt.Rows[0]["AppOverTime"]);
            }
            catch
            {
            }
            try
            {
                LossCard = Convert.ToInt32(dt.Rows[0]["LossCard"]);
            }
            catch
            {
            }
            
            

            if (Labsent != 0)//nghỉ phép
            {
                hlk_Wait_Labsent.Text = Resources.Strings.hlk_Wait_Labsent + Labsent.ToString() + Resources.Strings.petition;//signRpt
                hlk_Wait_Labsent.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitLabsentList.aspx";
            }

            if (OverTime != 0)//Tăng ca tính lương
            {
                hlk_Wait_OverTime.Text = Resources.Strings.hlk_Wait_OverTime + OverTime.ToString() + Resources.Strings.petition;//signRpt
                hlk_Wait_OverTime.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitOverTimeList.aspx";
            }


            if (MissCard != 0)//Quên bấm thẻ
            {
                hlk_Wait_MissCard.Text = Resources.Strings.hlk_Wait_MissCard + MissCard.ToString() + Resources.Strings.petition;//signRpt
                hlk_Wait_MissCard.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitMissCardList.aspx";
            }

            if (OverTime_Leave != 0)//Tăng ca nghỉ bù
            {
                hlk_Wait_OverTime_Leave.Text = Resources.Strings.hlk_Wait_OverTime_Leave + OverTime_Leave.ToString() + Resources.Strings.petition;//signRpt
                hlk_Wait_OverTime_Leave.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitOverTimeList_Leave.aspx";
            }
            if (AppOverTime != 0)//Tăng ca báo trước
            {
                hlk_Wait_OverTime_App.Text = Resources.Strings.hlk_Wait_OverTime_App + AppOverTime.ToString() + Resources.Strings.petition;
                hlk_Wait_OverTime_App.NavigateUrl = "~/PLSignWeb2/OverApply/WaitOverTime.aspx";
            }

            if (LossCard != 0)//Quên bấm thẻ do mất thẻ
            {
                hlk_Wait_LossCard.Text = Resources.Strings.hlk_Wait_LossCard + LossCard.ToString() + Resources.Strings.petition;//signRpt
                hlk_Wait_LossCard.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitLossCardList.aspx";
            }


            if (LevOut != 0)// Xin ra ngoai
            {
                hlk_wait_LeaveOut.Text = Resources.Strings.hlk_wait_lev_out + LevOut.ToString() + Resources.Strings.petition;
                hlk_wait_LeaveOut.NavigateUrl = "~/PLSignWeb2/WaitSignApplication/WaitLeaveOutList.aspx";
            }
        }
        
    }

    private DataTable GetCountData()
    {
        DataTable dt = new DataTable();
        PVN = new PersonnelVN();
        dt = PVN.Get_Count_Wait_List(strUserID, ApID);

        return dt;
    }

    private int GetWait_OverTime_App()
    {
        db_OverApplyGroup1 Over = new db_OverApplyGroup1(ConfigurationManager.AppSettings["PLVNConnectionType"], ConfigurationManager.AppSettings["PLVNConnectionServer"], ConfigurationManager.AppSettings["PLVNConnectionDB"], ConfigurationManager.AppSettings["PLVNConnectionUser"], ConfigurationManager.AppSettings["PLVNConnectionPwd"]);
        DataSet ds = Over.pro_get_wait_AppOvertime(ApID, strUserID, "0", "0", "0", "0", "", "");
        DataTable dt = new DataTable();
        dt = ds.Tables["waitOver_list"];
        return dt.Rows.Count;
    }

   
    //檢查授權是否正在啟用
    /*private int CheckAuthStatus()
    {
        int CountAuthRpt = 0;

        if (GetRptbyPLVN().Rows.Count != 0)
        {
            foreach (DataRow dr in GetRptbyPLVN().Rows)
            {
                DataTable dt = null;
                Flow myFl = new Flow();
                ArrayList AL = new ArrayList();

                AL.Add("rpt_id:" + dr["rpt_id"].ToString().Trim());
                AL.Add("user_id:" + strUserID);
                dt = myFl.SimpleGetData("w_authgroup", AL, "");
                if (dt.Rows.Count != 0)
                {
                    string end_mk = dt.Rows[0]["end_mk"].ToString();
                    if (end_mk == "N")//啟用
                    {
                        CountAuthRpt++;
                    }
                }

            }
        }

        return CountAuthRpt;
    }*/
    //檢查代理是否啟用
    /*private int ChkAgent()
    {
        int chkAgent = 0;

        if (GetRptbyPLVN().Rows.Count != 0)
        {
            foreach (DataRow dr in GetRptbyPLVN().Rows)
            {
                DataTable dt = null;
                Flow myFl = new Flow(); ;
                ArrayList AL = new ArrayList();
                AL.Add("rpt_id:" + dr["rpt_id"].ToString().Trim());
                AL.Add("user_id:" + strUserID);
                dt = myFl.SimpleGetData("w_agentd", AL, "");
                if (dt.Rows.Count != 0)
                {
                    string end_mk = dt.Rows[0]["end_mk"].ToString();

                    if (end_mk == "N")//啟用
                    {
                        chkAgent++;
                    }
                }

            }//foreach
        }

        return chkAgent;
    }*/
    /// <summary>
    /// 取得系統有幾張報表
    /// </summary>
    /*private DataTable GetRptbyPLVN()
    {
        DataTable dt = null;
        Flow fw = new Flow();
        HashMap hm = new HashMap();
        hm.Add("@ap_id", ApID);
        hm.Add("@agrp_nm", "");
        hm.Add("@agrp_id", 0);
        hm.Add("@counts", "OUTPUT");

        dt = fw.GetDataByProc("Pro_GetReportFromAuthGrp_Person", hm);

        return dt;
    }*/

    private string ConvertAreaName(string area_mk, string ap_id)
    {
        string strReturn = "";
        PccMsg myMsg = new PccMsg();
        //Steven:2009/03/13
        if (Session["PreferredCulture"].ToString() == "vi-VN")
            myMsg.Load(Server.MapPath(Session["PageLayer"] + MENUXML_VN));
        else
            myMsg.Load(Server.MapPath(Session["PageLayer"] + MENUXML));

        if (myMsg.QueryNodes("Applications/Application") != null)
        {
            foreach (XmlNode apNode in myMsg.QueryNodes("Applications/Application"))
            {
                if (myMsg.Query("ApID", apNode) == ap_id)
                {
                    if (myMsg.QueryNodes("ApAreas/Area", apNode) != null)
                    {
                        foreach (XmlNode areaNode in myMsg.QueryNodes("ApAreas/Area", apNode))
                        {
                            if (myMsg.Query("AreaMK", areaNode) == area_mk)
                                return myMsg.Query("AreaName", areaNode);
                        }
                    }
                    break;
                }
            }
        }
        return strReturn;
    }

    protected override void InitializeCulture()
    {
        // 用Session来存储语言信息
        if (Session["PreferredCulture"] == null)
            Session["PreferredCulture"] = Request.UserLanguages[0];
        string UserCulture = Session["PreferredCulture"].ToString();
        if (UserCulture != "")
        {
            //根据Session的值重新绑定语言代码
            Thread.CurrentThread.CurrentUICulture = new CultureInfo(UserCulture);
            Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture(UserCulture);
        }
    }
}

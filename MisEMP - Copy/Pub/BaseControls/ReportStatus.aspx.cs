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
using PLSignWeb2.Pub.Module;
using System.IO;
using System.Text;
using PccCommonForC;
using System.Net;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Xml;

public partial class PLSignWeb2_UserControl_ReportStatus : BasePage
{
    PersonnelVN PL;
    private string next_user = "";
    private bool m_status = false;
    private DataTable dt;
    protected void Page_Load(object sender, EventArgs e)
    {
        string FlowID = Commfun.CheckParams("FlowID", Page);
        string VouNO = Commfun.CheckParams("VouNO", Page);
        string Submit_id = Commfun.CheckParams("submit_id", Page);
        
        if (!IsPostBack)
        {
            if (FlowID != "")
            {
                ShowFlow(FlowID);
            }
            if (Submit_id != "") {
                ShowStatusBPMS(Submit_id);
            }
            else if (VouNO != "")
            {
                ShowStatus(VouNO);
            }
        }
    }

    #region " 檢視流程 "
    private void ShowFlow(string FlowID)
    {
        PL = new PersonnelVN();
        ArrayList AL = new ArrayList();
        AL.Add("flow_id : " + FlowID);
        DataTable dt = PL.SimpleGetData("vie_workflow", AL,"");
        
        if (dt.Rows.Count > 0)
        {
            string userstr = Commfun.CheckDBNull(dt.Rows[0]["flowstr"]);
            string[] userarray = userstr.Split(',');

            PccRow myRow;
            myRow = new PccRow("", HorizontalAlign.Center, 0, 0);

            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>申請人</font><div></td></tr></table>");
            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow01.gif'></td></tr></table>");
            ArrayList AL_user = new ArrayList();
            for (int i = 0; i < userarray.Length; i++)
            {
                AL_user.Clear();
                AL_user.Add("user_id :" + userarray[i].ToString());
                DataTable dt_user = PL.SimpleGetData("vie_employee", AL_user, "");
                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + Commfun.CheckDBNull(dt_user.Rows[0]["emp_name"]) + "<br>" + Commfun.CheckDBNull(dt_user.Rows[0]["posit_nm"]) + "</font><div></td></tr></table>");
               if (i != (userarray.Length -1))
                   myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow01.gif'></td></tr></table>");
               
            }
            mTable.Rows.Add(myRow.Row);
        }
    }
    #endregion

    #region " 目前流程狀態 "
    private void ShowStatus(string vou_no)
    {
        PL = new PersonnelVN();
        DataSet ds = PL.Get_Flow_Instance(vou_no);
        dt = ds.Tables[0];

        //next_user = ds.Tables[1].Rows[0]["next_user"].ToString();
        if (ds.Tables[1].Rows.Count>0) next_user = ds.Tables[1].Rows[0]["next_user"].ToString();
        //Steven:2009/03/28
        if (dt != null && dt.Rows.Count > 0)
        {

            DataRow[] drRoot = dt.Select("up_id = 0");

            PccRow myRow;
            myRow = new PccRow("", HorizontalAlign.Center, 0, 0);

            string root_user = drRoot[0]["user_id"].ToString();
            string rpt_id = drRoot[0]["rpt_id"].ToString();
            string fact_no = drRoot[0]["fact_no"].ToString();

            CreateTable(root_user,rpt_id, fact_no, myRow);
            if (!m_status)
                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drRoot[0]["emp_name"] + "<br>" + drRoot[0]["posit_nm"] +"<br>"+ drRoot[0]["email"]+ "</font><div></td></tr></table>");
            else
                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drRoot[0]["emp_name"] + "<br>" + drRoot[0]["posit_nm"] + "<br>" + drRoot[0]["email"] + "</font><font color=#e1e1f8>" + UserEmailAgent(root_user, rpt_id, fact_no) + "</font><div></td></tr></table>");
            mTable.Rows.Add(myRow.Row);
        }
        else
        {
            PccRow myRow;
            myRow = new PccRow("", HorizontalAlign.Center, 0, 0);
            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><font color=red>next user not exist</font></td></tr></table>");
            mTable.Rows.Add(myRow.Row);
        }


    }

    private void ShowStatusBPMS(string submit_id)
    {
       
        string BPMS_FlowSign = ConfigurationSettings.AppSettings["BPMS_FlowSign"];
        string BPMS_FlowSignLang = ConfigurationSettings.AppSettings["BPMS_FlowSignLang"];
        string url = BPMS_FlowSign + submit_id + BPMS_FlowSignLang;
        WebClient client = new WebClient();

        client.Encoding = System.Text.Encoding.UTF8;
        string json = client.DownloadString(url).ToString().Trim();

        byte[] utf8Bytes = Encoding.UTF8.GetBytes(json);
        string safeJsonStr = Encoding.UTF8.GetString(utf8Bytes);

        JObject results = JObject.Parse(safeJsonStr);
        
        int count = 0;
        int j = 0;

        JArray items = (JArray)results["result"];
        if (items != null)
        {
            PccRow myRow;
            myRow = new PccRow("", HorizontalAlign.Center, 0, 0);

            int length = items.Count;
            for (int i = 0; i < items.Count; i++)
            {
                count = count + 1;
                j = j + 1;
                string userName = (string)results["result"][i]["userName"];
                string userNameNext = "";
                string receiverDate = (string)results["result"][i]["receiverDate"];                
                string signDate = (string)results["result"][i]["signDate"];
                string sign_time = "";
                if (signDate.Length >= 12)
                    sign_time = signDate.Substring(0, 4) + "/" + signDate.Substring(4, 2) + "/" + signDate.Substring(6, 2) + " " + signDate.Substring(8, 2) + ":" + signDate.Substring(10, 2) + ":" + signDate.Substring(12, 2);
                string desc = (string)results["result"][i]["desc"];
                string positionName = (string)results["result"][i]["positionName"];

                if ((string)results["result"][i]["actionType"] == "PRX" && signDate != "")
                {
                    userNameNext = (string)results["result"][i]["userName"];
                    userName = (string)results["result"][i]["comment"];
                    userName = userName.Replace("Deputy of", "");
                    userName = userName.Replace("[", "");
                    userName = userName.Replace("]", "");
                }
                if (i + 1 < items.Count)
                {
                    if ((string)results["result"][i + 1]["actionType"] == "PRX" && signDate == "")
                    {
                        userNameNext = (string)results["result"][i + 1]["userName"];
                        i = i + 1;
                    }
                }
                
                string strTD = "";
                if (count > 1)
                {
                    strTD = "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow02.gif'><br><div style=\"padding:3;width:100%\"><div></td>";
                    if (signDate != "")
                    {
                        strTD = "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow01.gif'><br><div style=\"padding:3;width:100%\"><div></td>";
                    }
                    else
                    {
                        if (count != j)
                        {
                            strTD = "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow01.gif'><br><div style=\"padding:3;width:100%\"><div></td>";
                        }
                    }
                }

                if (signDate == "") j = 0;
                if (userNameNext == "")
                {                   
                    if (signDate != "")
                    {
                        myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "<br>" + sign_time + "</font><div></td></tr></table>");
                    }
                    else
                    {
                        myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "</font><div></td></tr></table>");
                        /*
                        if (receiverDate != "")
                        {
                            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "</font><div></td></tr></table>");
                        }
                        else
                        {
                            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "</font><div></td></tr></table>");
                        }
                        */
                    }
                }
                else
                {
                    if (signDate != "")
                    {
                        myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "<font color=#65a1dc>" + "<br>" + userNameNext + "<br>" + sign_time + "</font>" + "</font><div></td></tr></table>");
                    }
                    else
                    {
                        myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "<font color=#65a1dc>" + "<br>" + userNameNext + "</font>" + "</font><div></td></tr></table>");
                        /*
                        if (receiverDate != "") {
                            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "<font color=#65a1dc>" + "<br>" + userNameNext + "</font>" + "</font><div></td></tr></table>");
                        }
                        else
                        {
                            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr>" + strTD + "<td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + userName + "<font color=#65a1dc>" + "<br>" + userNameNext + "</font>" + "</font><div></td></tr></table>");
                        }
                        */
                    }
                }
                
                mTable.Rows.Add(myRow.Row);
            }

        }
        else
        {
            PccRow myRow;
            myRow = new PccRow("", HorizontalAlign.Center, 0, 0);
            myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><font color=red>BPMS 未接收單據, 請稍等. He thong BPMS chua nhan don, vui long cho trong it phut!</font></td></tr></table>");
            mTable.Rows.Add(myRow.Row);
        }
        

    }

    #endregion

    private string UserEmailAgent(string aUserID, string rpt_id, string sFact_no)
    {
        string s = "";
        DataTable dt = new DataTable();

        dt = PL.GetUserInfoAgent(aUserID, rpt_id, sFact_no).Tables[0];

        if (dt.Rows.Count > 0)
        {
            if (dt.Rows[0]["user_id_ag"].ToString() != "" && dt.Rows[0]["user_id_ag"].ToString() != null)
                s = "<b>(代)</b>" + "&nbsp;" + dt.Rows[0]["user_desc_ag"].ToString() + "<br>" + dt.Rows[0]["email_ag"].ToString();
        }
        return s;
    }

    private void CreateTable(string user_id,string rpt_id,string fact_no, PccRow myRow)
    {
        DataRow[] drChild = dt.Select("up_id = " + user_id);

        if (drChild.Length == 0)
            return;
        else
        {
            CreateTable(drChild[0]["user_id"].ToString(), rpt_id, fact_no, myRow);

            if (decimal.Parse(user_id.ToString()) == decimal.Parse(next_user.ToString()))
            {
                m_status = true;
                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drChild[0]["emp_name"] + "<br>" + drChild[0]["posit_nm"] + "<br>" + drChild[0]["email"] + "</font><div></td></tr></table>");
                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow02.gif'></td></tr></table>");
            }
            else
            {
                if (!m_status)
                    myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg02.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drChild[0]["emp_name"] + "<br>" + drChild[0]["posit_nm"] + "<br>" + drChild[0]["email"] + "</font><div></td></tr></table>"); 
                else
                    //myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drChild[0]["emp_name"] + "<br>" + drChild[0]["posit_nm"] + "<br>" + drChild[0]["email"] + "</font><div></td></tr></table>");
                    myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowImg01.gif'><br><div style=\"padding:3;width:100%\"><font size=2pt>" + drChild[0]["emp_name"] + "<br>" + drChild[0]["posit_nm"] + "<br>" + drChild[0]["email"] + "</font><font color=#e1e1f8>" + UserEmailAgent(drChild[0]["user_id"].ToString(), rpt_id, fact_no) + "</font><div></td></tr></table>");

                myRow.AddTextCell("<table border=\"0\" cellspacing=\"12\"><tr><td align=\"center\"><img src='../../App_Themes/pwfBody/images/reportImg/flowArrow01.gif'></td></tr></table>");
            }
        }
    }
}

using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Configuration;
using PccCommonForC;
using PccBsLayerForC;
using System.Collections.Generic; 

namespace PLSignWeb2.Pub.Module
{
	/// <summary>
	/// UserLogin 的摘要描述。
	/// </summary>
	public partial class UserLogin : System.Web.UI.Page
	{
        private Dictionary<string, string> KeyValue = HardCode.Hardcode_ByIDAndArea();
		protected void Page_Load(object sender, System.EventArgs e)
		{

			Hashtable myHT = new Hashtable(); 
				
			Session["UserName"] = "";
			Session["XmlLoginInfo"] = "";
			Session["APCounts"] = myHT;
			Session["UserIDAndName"] = Request.Params["REMOTE_ADDR"];

			Session["PageLayer"] = "../../";
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 此呼叫為 ASP.NET Web Form 設計工具的必要項。
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 此為設計工具支援所必需的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{    

		}
		#endregion

		protected void btnClear_Click(object sender, System.EventArgs e)
		{
			txtUserName.Text = "";
			txtPassWord.Text = "";
		}

		protected void btnLogin_Click(object sender, System.EventArgs e)
		{
			PccMsg myMsg = new PccMsg("","Big5");
			bs_Security mySecurity = new bs_Security(ConfigurationSettings.AppSettings["ConnectionType"] , ConfigurationSettings.AppSettings["ConnectionServer"], ConfigurationSettings.AppSettings["ConnectionDB"], ConfigurationSettings.AppSettings["ConnectionUser"], ConfigurationSettings.AppSettings["ConnectionPwd"],Session["UserIDAndName"].ToString(),ConfigurationSettings.AppSettings["EventLogPath"]);
			PccErrMsg myErrMsg = new PccErrMsg(Server.MapPath(Session["PageLayer"] + "XmlDoc"),Application["CodePage"].ToString() ,"Error");
			string strXmlReturn;

			myMsg.CreateFirstNode("UserName",txtUserName.Text);
			myMsg.CreateFirstNode("Password",txtPassWord.Text);

			myMsg.CreateFirstNode("vpath",ConfigurationSettings.AppSettings["vpath"]);
			myMsg.CreateFirstNode("superAdmin",KeyValue["superAdminEmail"]);

			strXmlReturn = mySecurity.DoReturnStr("GetUserInfo",myMsg.GetXmlStr,"");

			myMsg.LoadXml(strXmlReturn);

			if (myMsg.Query("Exist") == "Y")
			{
				Session["XmlLoginInfo"] = strXmlReturn;
				Session["UserName"] = myMsg.Query("UserDesc");
				Session["UserAccount"] = myMsg.Query("UserName");
				Session["UserID"] = myMsg.Query("UserID");
				Session["UserIDAndName"] = myMsg.Query("UserID") + "---" + myMsg.Query("UserDesc") + "---" + Request.Params["REMOTE_ADDR"];

				string FilePath = ConfigurationSettings.AppSettings["myServer"].ToString() + ConfigurationSettings.AppSettings["vpath"].ToString();
				GetMenuAuth myAuth = new GetMenuAuth();
				if (myAuth.IsApAuth())
				{
					lblErrorMsg.Text ="";
					switch (CheckQueryString("ApID"))
					{
						case "104":
							FilePath += "/SysManager/SysManagerHome.aspx?ApID=104";
							break;
						default:
							FilePath += "/Pub/Module/pgError.aspx?ApID=" + CheckQueryString("ApID");
							break;
					}
					Response.Redirect(FilePath); 
				}
				else
				{
					lblErrorMsg.Text = myErrMsg.GetErrMsg("msg0054");
				}
			}
			else
			{
				lblErrorMsg.Text = myErrMsg.GetErrMsg("msg0030");
			}
		
		}

		#region "Tool Func. ex. CheckDBNull()"
		
		private string CheckDBNull(object oFieldData)
		{
			if (Convert.IsDBNull(oFieldData))
				return "";
			else
				return oFieldData.ToString(); 
		}

		private string CheckForm(string strName)
		{
			if (Request.Form[strName] == null)
				return "";
			else
				return Request.Form[strName].ToString();
		}

		private string CheckQueryString(string strName)
		{
			if (Request.QueryString[strName] == null)
				return "";
			else
				return Request.QueryString[strName].ToString();
		}

		private string CheckParams(string strName)
		{
			if (Request.Params[strName] == null)
				return "";
			else
				return Request.Params[strName].ToString();
		}

		#endregion

	}
}

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

namespace PLSignWeb2.Pub.Module
{
	/// <summary>
	/// ChangeLanguage 的摘要描述。
	/// </summary>
	public partial class ChangeLanguage : System.Web.UI.Page
	{
	
		protected void Page_Load(object sender, System.EventArgs e)
		{
			// 將使用者程式碼置於此以初始化網頁
			if (Request.QueryString["ApID"] != null && Request.QueryString["ApID"] == "")
			{
				DefaultPage();
			}
			else
			{
				switch (Request.QueryString["ApID"].ToString())
				{
					case "0":
						DefaultPage();
						break;
					case "A":
						AboutBody();
						break;
					case "104":
						SysManageBody();
						break;
				}
			}
		}

		private void DefaultPage()
		{
			//新增PccApHome中所要顯示的元素 20040525

			PccRow myRow = new PccRow();
			myRow.AddTextCell("&nbsp;&nbsp;&nbsp;<font size=large color=blue><b>" + Session["UserName"] + "</b>Welcome to Pcc Ap Home</font>",100); 
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset();
			myRow.AddTextCell("<br><font color=red><b>如您需修改個人資料,可點選左上方第一個分區點選個人資料修改選單，以進行個人資料修改</b></font>",100); 
			tblBody.Rows.Add(myRow.Row);
			
			//設定中英文轉換之Button
			if (Application["CodePage"].ToString() == "CP950")
				LinkButton1.Text = "Do you want transfer to English?";
			else
				LinkButton1.Text = "您想要轉換到中文嗎?";
     

			LinkButton1.Visible = true;
		}

		private void AboutBody()
		{
			PccRow myRow = new PccRow();
			myRow.AddTextCell("&nbsp;&nbsp;&nbsp;<font size=large color=blue><b>寶成國際集團系統網頁</b></font>",100);
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset(); 
			myRow.AddTextCell("<br><font color=red>所有的權利屬於<font color=blue><b>寶成國際集團</b></font></font>",100);
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset();
			myRow.AddTextCell("<br><font color=blue>版本：1.0</font>",100);
			tblBody.Rows.Add(myRow.Row);
		}

		private void SysManageBody()
		{
			//在這裡可以寫入進入這個系統中所要顯示的首頁資訊，或是
			//由UserLogin.aspx中傳入之Params，利用其XML傳入的
			//參數做判斷以使這頁導引至正確的網頁。	
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

		protected void LinkButton1_Click(object sender, System.EventArgs e)
		{
			//設定中英文轉換之Button
			if (Application["CodePage"].ToString()  == "CP950")
			{
				Application["CodePage"] = "CP437";
				LinkButton1.Text = "您想要轉換到中文嗎?";
			}
			else
			{
				Application["CodePage"] = "CP950";
				LinkButton1.Text = "Do you want transfer to English?";
			}
		}
	}
}

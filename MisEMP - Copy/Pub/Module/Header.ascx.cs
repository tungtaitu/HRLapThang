namespace PLSignWeb2.Pub.Module
{
	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;

	/// <summary>
	///		Header 的摘要描述。
	/// </summary>
	public partial  class Header : System.Web.UI.UserControl
	{


		protected void Page_Load(object sender, System.EventArgs e)
		{
			if (Session["UserID"] == null) return;
			
			// 將使用者程式碼置於此以初始化網頁
			Pub.Module.GetMenuAuth myAuth = new Pub.Module.GetMenuAuth();
			string menu_nm = myAuth.GetMenu();
			//判斷是否為空白則就由程式員自己寫入的Title來呈現，否則就直接用Menu_nm的字串。 20040607
			if (menu_nm != "")
			{
				if (PccTitle.Text.Equals(string.Empty) || PccTitle.Text.Equals("Title"))
					PccTitle.Text = myAuth.GetMenu();
			}
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 此呼叫為 ASP.NET Web Form 設計工具的必要項。
			//
			
			//計算這個網頁所在的層次是那裡
			int i,j = 0;
			string strPageLayer = "";
			string LocalPath = PccCommonForC.PccToolFunc.Upper(Server.MapPath("."));   

			j = LocalPath.IndexOf(PccCommonForC.PccToolFunc.Upper(Application["EDPNET"].ToString()));
		
			try
			{
				for (i = 1 ; i < LocalPath.Substring(j).Split('\\').Length ; i++)
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
		/// 此為設計工具支援所必需的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{
		}
		#endregion
	}
}

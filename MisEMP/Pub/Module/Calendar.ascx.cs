	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;

	/// <summary>
	///		Calendar 的摘要描述。
	/// </summary>
    public partial class Pub_CommControl_Calendar : System.Web.UI.UserControl
	{

		protected void Page_Load(object sender, System.EventArgs e)
		{
			// 將使用者程式碼置於此以初始化網頁
		}

		public string Text
		{
			get
			{
				return UC_Calendar.Text;
			}
			set
			{
				UC_Calendar.Text = value;
			}
		}

		public bool Enabled
		{
			get
			{
				return UC_Calendar.Enabled;
			}
			set
			{
				UC_Calendar.Enabled = value;
			}
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
	}

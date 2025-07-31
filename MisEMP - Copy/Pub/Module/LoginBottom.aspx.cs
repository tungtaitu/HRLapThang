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
using System.Collections.Generic;

namespace PLSignWeb2
{
	/// <summary>
	/// Bottom 的摘要描述。
	/// </summary>
	public partial class Bottom : System.Web.UI.Page
	{
        public Dictionary<string, string> KeyValue = HardCode.Hardcode_ByIDAndArea();
		protected void Page_Load(object sender, System.EventArgs e)
		{
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
}

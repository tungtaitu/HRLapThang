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
	/// pgError ���K�n�y�z�C
	/// </summary>
	public partial class pgError : System.Web.UI.Page
	{
	
		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			string ServerPath = System.Configuration.ConfigurationSettings.AppSettings["myServer"].ToString();
			reLogin.NavigateUrl = ServerPath + System.Configuration.ConfigurationSettings.AppSettings["vpath"].ToString()  + "/Default.aspx";
		}

		#region Web Form Designer generated code
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: ���I�s�� ASP.NET Web Form �]�p�u�㪺���n���C
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// �����]�p�u��䴩�ҥ��ݪ���k - �ФŨϥε{���X�s�边�ק�
		/// �o�Ӥ�k�����e�C
		/// </summary>
		private void InitializeComponent()
		{    

		}
		#endregion
	}
}

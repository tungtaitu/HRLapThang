namespace PLSignWeb2.Pub.Module
{
	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;

	/// <summary>
	///		Header ���K�n�y�z�C
	/// </summary>
	public partial  class Header : System.Web.UI.UserControl
	{


		protected void Page_Load(object sender, System.EventArgs e)
		{
			if (Session["UserID"] == null) return;
			
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			Pub.Module.GetMenuAuth myAuth = new Pub.Module.GetMenuAuth();
			string menu_nm = myAuth.GetMenu();
			//�P�_�O�_���ťիh�N�ѵ{�����ۤv�g�J��Title�ӧe�{�A�_�h�N������Menu_nm���r��C 20040607
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
			// CODEGEN: ���I�s�� ASP.NET Web Form �]�p�u�㪺���n���C
			//
			
			//�p��o�Ӻ����Ҧb���h���O����
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
		/// �����]�p�u��䴩�ҥ��ݪ���k - �ФŨϥε{���X�s�边�ק�
		/// �o�Ӥ�k�����e�C
		/// </summary>
		private void InitializeComponent()
		{
		}
		#endregion
	}
}

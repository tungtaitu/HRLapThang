	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;

	/// <summary>
	///		Calendar ���K�n�y�z�C
	/// </summary>
    public partial class Pub_CommControl_Calendar : System.Web.UI.UserControl
	{

		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
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

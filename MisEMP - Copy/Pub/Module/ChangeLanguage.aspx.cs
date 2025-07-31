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
	/// ChangeLanguage ���K�n�y�z�C
	/// </summary>
	public partial class ChangeLanguage : System.Web.UI.Page
	{
	
		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
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
			//�s�WPccApHome���ҭn��ܪ����� 20040525

			PccRow myRow = new PccRow();
			myRow.AddTextCell("&nbsp;&nbsp;&nbsp;<font size=large color=blue><b>" + Session["UserName"] + "</b>Welcome to Pcc Ap Home</font>",100); 
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset();
			myRow.AddTextCell("<br><font color=red><b>�p�z�ݭק�ӤH���,�i�I�索�W��Ĥ@�Ӥ����I��ӤH��ƭק���A�H�i��ӤH��ƭק�</b></font>",100); 
			tblBody.Rows.Add(myRow.Row);
			
			//�]�w���^���ഫ��Button
			if (Application["CodePage"].ToString() == "CP950")
				LinkButton1.Text = "Do you want transfer to English?";
			else
				LinkButton1.Text = "�z�Q�n�ഫ�줤���?";
     

			LinkButton1.Visible = true;
		}

		private void AboutBody()
		{
			PccRow myRow = new PccRow();
			myRow.AddTextCell("&nbsp;&nbsp;&nbsp;<font size=large color=blue><b>�_����ڶ��Ψt�κ���</b></font>",100);
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset(); 
			myRow.AddTextCell("<br><font color=red>�Ҧ����v�Q�ݩ�<font color=blue><b>�_����ڶ���</b></font></font>",100);
			tblBody.Rows.Add(myRow.Row);
			myRow.Reset();
			myRow.AddTextCell("<br><font color=blue>�����G1.0</font>",100);
			tblBody.Rows.Add(myRow.Row);
		}

		private void SysManageBody()
		{
			//�b�o�̥i�H�g�J�i�J�o�Өt�Τ��ҭn��ܪ�������T�A�άO
			//��UserLogin.aspx���ǤJ��Params�A�Q�Ψ�XML�ǤJ��
			//�Ѽư��P�_�H�ϳo���ɤަܥ��T�������C	
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

		protected void LinkButton1_Click(object sender, System.EventArgs e)
		{
			//�]�w���^���ഫ��Button
			if (Application["CodePage"].ToString()  == "CP950")
			{
				Application["CodePage"] = "CP437";
				LinkButton1.Text = "�z�Q�n�ഫ�줤���?";
			}
			else
			{
				Application["CodePage"] = "CP950";
				LinkButton1.Text = "Do you want transfer to English?";
			}
		}
	}
}

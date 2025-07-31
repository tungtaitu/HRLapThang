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
using System.Xml;
using PccBsLayerForC;
using System.Configuration; 


namespace PLSignWeb2.Pub.Module
{
	/// <summary>
	/// Medium1 ���K�n�y�z�C
	/// </summary>
	public partial class Medium1 : System.Web.UI.Page
	{
	
		/*var menus = new Array(
				2,	
							
				"<image src=../../images/MenuArea/N_Close.gif />",		
							
				1,
				"���ε{���޲z",
				"../../images/MenuArea/DgyyWebWinNew/sFile1.gif",
				"../../GenCodeAP/APManage/APManage141.aspx?ApID=141",	
				"1",					
							
				"<image src=../../images/MenuArea/Y_Close.gif />",										
							
				2,					
				"�s�պ޲z",									
				"../../images/MenuArea/DgyyWebWinNew/sFile1.gif",						
				"../../SysManager/GroupManage/GroupManage104.aspx?ApID=104",				
				"1",												
							
				"�ϥΪ̺޲z",										
				"../../images/MenuArea/DgyyWebWinNew/sQuery1.gif",				
				"../../SysManager/UserManage/UserManage104.aspx?ApID=104",				
				"2"												
				);
				
				//�}�l�g�J
				strLeftMenu += "var menus = new Array (";
				strLeftMenu += "2,"; //�`�@�h�֤���
					
				strLeftMenu += "\"<image src=../../images/MenuArea/N_Open.gif />\","; //�Ĥ@�Ӥ��Ϫ��ϧΩΤ�r
				strLeftMenu += "1,"; //�Ĥ@�Ӥj�������}�l�A��ܦ��h�֭�Item
				//�Ĥ@�Ӥ��Ϫ��Ĥ@�ӤpItem�Ѽ�
				strLeftMenu += "\"���ε{���޲z\","; //���W��
				strLeftMenu += "\"../../images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
				strLeftMenu += "\"../../GenCodeAP/APManage/APManage141.aspx?ApID=141\",";//��檺�s������
				strLeftMenu += "\"1\","; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C

				strLeftMenu += "\"<image src=../../images/MenuArea/Y_Open.gif />\","; //�ĤG�Ӥ��Ϫ��ϧΩΤ�r
				strLeftMenu += "2,"; //�ĤG�Ӥj�������}�l�A��ܦ��h�֭�Item
				//�ĤG�Ӥ��Ϫ��Ĥ@�ӤpItem�Ѽ�
				strLeftMenu += "\"�s�պ޲z\","; //���W��
				strLeftMenu += "\"../../images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
				strLeftMenu += "\"../../SysManager/GroupManage/GroupManage104.aspx?ApID=104\",";//��檺�s������
				strLeftMenu += "\"1\","; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C
				//�ĤG�Ӥ��Ϫ��ĤG�ӤpItem�Ѽ�
				strLeftMenu += "\"�ϥΪ̺޲z\","; //���W��
				strLeftMenu += "\"../../images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
				strLeftMenu += "\"../../SysManager/UserManage/UserManage104.aspx?ApID=104\",";//��檺�s������
				strLeftMenu += "\"1\""; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C

				strLeftMenu += ");";*/

		private Hashtable m_Menu = new Hashtable(); 

		protected void Page_Load(object sender, System.EventArgs e)
		{
			if (Session["UserID"] == null) return;
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			string strLeftMenu = string.Empty;

			int AreaCount = 1;

			//�Q��LoginInfo���o��User��AP Detail Menu.
			PccCommonForC.PccErrMsg myLabel = new PccCommonForC.PccErrMsg(Server.MapPath(Session["PageLayer"] + "XmlDoc"),Application["CodePage"].ToString() ,"Label");
			string strDetailPageLayer = Session["PageLayer"].ToString();

			


			if (Request.QueryString["ApID"] != null)
			{
				//���o���ε{����APID
				string strApID = Request.Params["ApID"];
				string ApFolder = "";

				PccCommonForC.PccMsg myMsg = new PccCommonForC.PccMsg(); 
				myMsg.LoadXml(Session["XmlLoginInfo"].ToString());

				if (strApID.Length > 0 && strDetailPageLayer.Length > 0)
				{
					//strDetailPageLayer = strDetailPageLayer.Substring(0,strDetailPageLayer.Length - 3);

					//�P�_�O�_���t�Υi�H�ϥΡA�Y�S���h�������X
					if (myMsg.QueryNodes("Authorize") != null)
					{
						foreach (XmlNode myNode in myMsg.QueryNodes("Authorize"))
						{
							//���p���O�I��o�Өt�Ϋh�������X�çP�_�U�@��
							if (strApID != myMsg.Query("APID",myNode)) continue; 

							//���o�o��AP���̤W�h��Folder
							ApFolder = GetApFolder(myMsg.Query("APLink",myNode)); 
							
							//�P�_�O�_�����t�Υi�H�ϥΡA�Y�S���h�������X
							if (myMsg.QueryNodes("ApMenu",(XmlElement)myNode) != null)
							{
								foreach(XmlNode myDetailNode in myMsg.QueryNodes("ApMenu",(XmlElement)myNode))
								{
									if (myMsg.Query("show_mk",myDetailNode).Equals("Y")) 
									{
										//�P�_�O�_��Web Service��Menu�Y�O�h�������L 2005/3/8
										if (myMsg.Query("MenuLink",myDetailNode).IndexOf(".asmx",0) > 0)
											continue;

										SaveDataToHashMenu(myMsg.Query("MenuManage",(XmlElement)myDetailNode),strDetailPageLayer,ApFolder,myMsg.Query("MenuNM",(XmlElement)myDetailNode),myMsg.Query("MenuLink",(XmlElement)myDetailNode));
									}
								}
							}
						}
					}
				}

				strLeftMenu += "var menus = new Array (";
				int count = 1;
				
				if (m_Menu.Count > 0)
				{
					AreaCount += m_Menu.Count;
					strLeftMenu += AreaCount.ToString()  + ","; //�`�@�h�֤���
					strLeftMenu += GetWelcome(strDetailPageLayer,strApID) + ",";
					
					//�]�w�@���
					if (m_Menu.ContainsKey("N"))
					{
						strLeftMenu += "\"<image src=" + strDetailPageLayer + "images/MenuArea/N_Open.gif />\",";
						count = m_Menu["N"].ToString().Split(',').Length / 4; 
						strLeftMenu += count.ToString()  + ","; //�@�Ӥj�������}�l�A��ܦ��h�֭�Item
						strLeftMenu += m_Menu["N"].ToString(); 
					}
					
					foreach (string strArea in m_Menu.Keys)
					{
						if (!strArea.Equals("N") && !strArea.Equals("Y"))
						{
							strLeftMenu += "\"<image src=" + strDetailPageLayer + "images/MenuArea/" + strArea + "_Open.gif />\",";
							count = m_Menu[strArea].ToString().Split(',').Length / 4; 
							strLeftMenu += count.ToString()  + ","; //�Ĥ@�Ӥj�������}�l�A��ܦ��h�֭�Item
							strLeftMenu += m_Menu[strArea].ToString(); 
						}
					}

					//�]�w�v���޲z��
					if (m_Menu.ContainsKey("Y"))
					{
						strLeftMenu += "\"<image src=" + strDetailPageLayer + "images/MenuArea/Y_Open.gif />\",";
						count = m_Menu["Y"].ToString().Split(',').Length / 4; 
						strLeftMenu += count.ToString()  + ","; //�@�Ӥj�������}�l�A��ܦ��h�֭�Item
						strLeftMenu += m_Menu["Y"].ToString(); 
					}

					strLeftMenu = strLeftMenu.Substring(0,strLeftMenu.Length - 1); 
				}
				else //��ܨS������v��
				{
					strLeftMenu += AreaCount.ToString()  + ","; //�`�@�h�֤���
					strLeftMenu += GetWelcome(strDetailPageLayer,strApID);
				}

				
				strLeftMenu += ");";
				
				
			}
			
			txtLeftMenu.Value = strLeftMenu; 

		}

		private void SaveDataToHashMenu(string Area,string strLayer,string ApFolder,string menuNm,string menuLink)
		{
			string strReturn = string.Empty;
 
			if (m_Menu.ContainsKey(Area)) 
				strReturn = m_Menu[Area].ToString();
 
			strReturn += "\"" + menuNm + "\","; //���W��
			strReturn += "\"" + strLayer + "images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
			//�P�_�O�_���v���޲z��
			if (Area.Equals("Y"))
				strReturn += "\"" + strLayer + "SysManager/" + menuLink + "\",";//��檺�s������
			else
				strReturn += "\"" + strLayer + ApFolder + "/" + menuLink + "\",";//��檺�s������

			strReturn += "\"1\","; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C
			
			if (m_Menu.ContainsKey(Area)) 
				m_Menu[Area] = strReturn; 
			else
				m_Menu.Add(Area,strReturn); 
			
		}

		private string GetApFolder(string ApLink)
		{
			/*int pos = ApLink.IndexOf("/");
			string strReturn = ApLink.Substring(0,pos);
			return strReturn;*/

			int pos = ApLink.LastIndexOf("&");
			string strReturn1 = ApLink.Substring(pos + 1);
			string strReturn = strReturn1.Split('=')[1].Trim();
			return strReturn;
		}

		private string GetWelcome(string strLayer,string strApID)
		{
			string strReturn = string.Empty;
			
			bs_Security mySecurity = new bs_Security(ConfigurationSettings.AppSettings["ConnectionType"] , ConfigurationSettings.AppSettings["ConnectionServer"], ConfigurationSettings.AppSettings["ConnectionDB"], ConfigurationSettings.AppSettings["ConnectionUser"], ConfigurationSettings.AppSettings["ConnectionPwd"],Session["UserIDAndName"].ToString(),ConfigurationSettings.AppSettings["EventLogPath"]);
			string strCount = "0";
			PccCommonForC.PccMsg myMsg1 = new PccCommonForC.PccMsg(); 

			if (strApID != null && int.Parse(strApID) > 0)
			{
				//�s�W�o�Өt�Ϊ��e�m��
				myMsg1.CreateFirstNode("ap_id",strApID); 
				myMsg1.CreateFirstNode("user_id",Session["UserID"].ToString());

				if (((Hashtable)Session["APCounts"]).ContainsKey(strApID))
				{
					strCount = ((Hashtable)Session["APCounts"])[strApID].ToString(); 
				}
				else
				{
					strCount = mySecurity.DoReturnStr("GetAndUpdateApCounts",myMsg1.GetXmlStr,""); 
					((Hashtable)Session["APCounts"]).Add(strApID,strCount); 
				}
			}

			strReturn += "\"�w��" + Session["UserName"].ToString() + "���{(" + strCount + ")\","; //�Ĥ@�Ӥ��Ϫ��ϧΩΤ�r
			strReturn += "3,"; //�Ĥ@�Ӥj�������}�l
			//�Ĥ@�Ӥ��Ϫ��Ĥ@�ӤpItem�Ѽ�
			strReturn += "\"�ӤH��ƭק�\","; //���W��
			strReturn += "\"" + strLayer + "images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
			strReturn += "\"" + strLayer + "UpdateLoginUser.aspx\",";//��檺�s������
			strReturn += "\"1\","; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C

			//�Ĥ@�Ӥ��Ϫ��ĤG�ӤpItem�Ѽ�
			strReturn += "\"�[�J�t��\","; //���W��
			strReturn += "\"" + strLayer + "images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
			strReturn += "\"ApplyAccount.aspx?Type=Update\",";//��檺�s������
			strReturn += "\"1\","; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C

			//�Ĥ@�Ӥ��Ϫ��ĤT�ӤpItem�Ѽ�
			strReturn += "\"���^���ഫ\","; //���W��
			strReturn += "\"" + strLayer + "images/MenuArea/DgyyWebWinNew/sFile1.gif\","; //��檺�e�m�ϧ�
			strReturn += "\"ChangeLanguage.aspx?ApID=0\",";//��檺�s������
			strReturn += "\"1\""; //���A��ܶ}�Ҥ@�ӭ����bIFrame�W�A�Y��2��ܩI�s�@�Ө禡�C

			return strReturn;
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

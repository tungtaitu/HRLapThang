namespace PLSignWeb2.Pub.Module
{
	using System;
	using System.Data;
	using System.Drawing;
	using System.Web;
	using System.Web.UI.WebControls;
	using System.Web.UI.HtmlControls;
	using System.Xml; 

	/// <summary>
	///		Menus ���K�n�y�z�C
	/// </summary>
	public partial  class Menus : System.Web.UI.UserControl
	{


		protected void Page_Load(object sender, System.EventArgs e)
		{
			// �N�ϥΪ̵{���X�m�󦹥H��l�ƺ���
			if (Session["UserName"] != null && Session["UserName"].ToString()  != "")
			{
				//���o���ε{����APID
				string strApID = Request.Params["ApID"];
				string ApFolder = "";
				
				if (strApID == null)
					strApID = "";

				//TableRow Rows;
				TableCell Cells;

				//�ϱo�Ĥ@�C�PMenu�W���@�ǶZ��
				PccRow myRow = new PccRow(); 
				myRow.AddTextCell(" ",0);
				tblSpace.Rows.Add(myRow.Row);  
				
				//���o��User�ү�ϥΤ�AP�A�ʺA���ͥD�n��Menu
				int counts = 0;
				PccCommonForC.PccMsg myMsg = new PccCommonForC.PccMsg(); 

				tblMainMenu.CssClass = "topmenu";
				Session["SupAdmin"] = "N";

				myMsg.LoadXml(Session["XmlLoginInfo"].ToString());

				//�P�_�O�_���t�Υi�H�ϥΡA�Y�S���h�������X
				if (myMsg.QueryNodes("Authorize") != null)
				{ 
					myRow.Reset(); 
					//�����[�J�W�h��Ap Menu�A�C�W�L���ӴN�[�J�@�C 2004/3/4
					foreach(XmlNode myNode in myMsg.QueryNodes("Authorize"))
					{
						counts += 1;
						myRow.AddTextCell("<A href=" + Session["PageLayer"] + myMsg.Query("APLink",(XmlElement)myNode) + ">" + myMsg.Query("APName",(XmlElement)myNode) + "</A>" ,0); 
						if ((Request.Params["ApID"] != null) && (myMsg.Query("APID") == Request.Params["ApID"].ToString()))
							Session["SupAdmin"] = myMsg.Query("SupAdmin",(XmlElement)myNode); 
						if (counts % 5 == 0)
						{
							tblMainMenu.Rows.Add(myRow.Row);
							myRow.Reset(); 
						}
					}

					if (counts % 5 != 0)
						tblMainMenu.Rows.Add(myRow.Row);  
				}	

				//�Q��LoginInfo���o��User��AP Detail Menu.
				PccCommonForC.PccErrMsg myLabel = new PccCommonForC.PccErrMsg(Server.MapPath(Session["PageLayer"] + "XmlDoc"),Application["CodePage"].ToString() ,"Label");
				string strDetailPageLayer = Session["PageLayer"].ToString();
				
				Table tblGeneror =  new Table();
				Table tblManage =  new Table();

				Table tblGenerorA =  new Table();
				Table tblGenerorB =  new Table();
				Table tblGenerorC =  new Table();
				Table tblGenerorD =  new Table();
				Table tblGenerorE =  new Table();
				Table tblGenerorF =  new Table();
				Table tblGenerorG =  new Table();
				Table tblGenerorH =  new Table();
				Table tblGenerorI =  new Table();
            
				//�ϱo�Ĥ@�C�P�ϧη|���@�w���Z��
				myRow.Reset(); 
				Cells = new TableCell();
				Cells.Height = Unit.Pixel(20);
				myRow.Row.Cells.Add(Cells); 
				tblDetailMenu.Rows.Add(myRow.Row);
				
				//�[�J�w�諸�e�m���A�����ϥΪ̥i�HLink��UpdateLoginUser���������ӭק�ۤv�����
				myRow.Reset(); 
				myRow.SetDefaultCellData("CellWelcome",HorizontalAlign.Left,VerticalAlign.Middle,0);
				myRow.AddTextCell(myLabel.GetErrMsg("msgWelcome") + "<A href=" + strDetailPageLayer + "UpdateLoginUser.aspx>" + Session["UserName"] + "</A>" + myLabel.GetErrMsg("msgWelcome2"),0); 
				tblDetailMenu.Rows.Add(myRow.Row);
				
				if (strApID.Length > 0 && strDetailPageLayer.Length > 0)
				{
					strDetailPageLayer = strDetailPageLayer.Substring(0,strDetailPageLayer.Length - 3);

					//�P�_�O�_���t�Υi�H�ϥΡA�Y�S���h�������X
					if (myMsg.QueryNodes("Authorize") != null)
					{
						foreach (XmlNode myNode in myMsg.QueryNodes("Authorize"))
						{
							//���p���O�I��o�Өt�Ϋh�������X�çP�_�U�@��
							if (strApID != myMsg.Query("APID",myNode)) continue; 

							//���o�o��AP���̤W�h��Folder
							ApFolder = GetApFolder(myMsg.Query("APLink",myNode)); 

							//�s�W�o�Өt�Ϊ��e�m��
							myRow.Reset();
							myRow.SetDefaultCellData("menuheading",0,0,2);
							myRow.AddTextCell(myMsg.Query("APName",(XmlElement)myNode) + myLabel.GetErrMsg("msgMenu"),0); 
							tblDetailMenu.Controls.Add(myRow.Row);

							//�P�_�O�_���t�Υi�H�ϥΡA�Y�S���h�������X
							if (myMsg.QueryNodes("ApMenu",(XmlElement)myNode) != null)
							{
								foreach(XmlNode myDetailNode in myMsg.QueryNodes("ApMenu",(XmlElement)myNode))
								{
									if (myMsg.Query("show_mk",myDetailNode) == "Y") 
									{
										myRow.Reset();
										myRow.SetDefaultCellData("menuitem",0,0,0);
										//�P�_�Y�O�޲z�ϫh�����s����SysManager��Floder��
										if (myMsg.Query("MenuManage",(XmlElement)myDetailNode) == "Y")
										{
											myRow.AddTextCell("<A href=" + "../" + strDetailPageLayer + "SysManager/" + myMsg.Query("MenuLink",(XmlElement)myDetailNode) + " >" + myMsg.Query("MenuNM",(XmlElement)myDetailNode) + "</A>",90);
										}
										else
										{
											myRow.AddTextCell("<A href=" + "../" + strDetailPageLayer + ApFolder + "/" + myMsg.Query("MenuLink",(XmlElement)myDetailNode) + " >" + myMsg.Query("MenuNM",(XmlElement)myDetailNode) + "</A>",90);
										}
										myRow.AddTextCell(" ",10); 
										//�P�_Menu�����ϡA�å[�J���j�Ÿ� 20040108
										switch (myMsg.Query("MenuManage",(XmlElement)myDetailNode))
										{
											case "A":
												tblGenerorA.Controls.Add(myRow.Row);
												break;
											case "B":
												tblGenerorB.Controls.Add(myRow.Row);
												break;
											case "C":
												tblGenerorC.Controls.Add(myRow.Row);
												break;
											case "D":
												tblGenerorD.Controls.Add(myRow.Row);
												break;
											case "E":
												tblGenerorE.Controls.Add(myRow.Row);
												break;
											case "F":
												tblGenerorF.Controls.Add(myRow.Row);
												break;
											case "G":
												tblGenerorG.Controls.Add(myRow.Row);
												break;
											case "H":
												tblGenerorH.Controls.Add(myRow.Row);
												break;
											case "I":
												tblGenerorI.Controls.Add(myRow.Row);
												break;
											case "N":
												tblGeneror.Controls.Add(myRow.Row);
												break;
											case "Y":
												tblManage.Controls.Add(myRow.Row);
												break;
											default:
												tblManage.Controls.Add(myRow.Row);
												break;
										} // end switch
									}//end check the show_mk
								} //end foreach Detail Menu
								CheckAndAddToTable(ref tblDetailMenu,ref tblGeneror,myLabel.GetErrMsg("ViewGeneror"),"doSectionMenu(Menus1_view_Generor,this)","N_Open.gif","on","view_Generor");
								CheckAndAddToTable(ref tblDetailMenu,ref tblManage,myLabel.GetErrMsg("ViewManage"),"doSectionOtherMenu(Menus1_view_Manage,this,'Y')","Y_Close.gif","off","view_Manage");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorA,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorA,this,'A')","A_Close.gif","off","view_GenerorA");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorB,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorB,this,'B')","I_Close.gif","off","view_GenerorB");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorC,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorC,this,'C')","I_Close.gif","off","view_GenerorC");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorD,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorD,this,'D')","I_Close.gif","off","view_GenerorD");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorE,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorE,this,'E')","I_Close.gif","off","view_GenerorE");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorF,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorF,this,'F')","I_Close.gif","off","view_GenerorF");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorG,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorG,this,'G')","I_Close.gif","off","view_GenerorG");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorH,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorH,this,'H')","I_Close.gif","off","view_GenerorH");
								CheckAndAddToTable(ref tblDetailMenu,ref tblGenerorI,myLabel.GetErrMsg("ViewNormal"),"doSectionOtherMenu(Menus1_view_GenerorI,this,'I')","I_Close.gif","off","view_GenerorI");
							}
						}//end foreach Ap myNode
					}
				} //end check ap_id
			} //end if username is space.
			else
			{
				Response.Redirect(Session["PageLayer"] + "Default.aspx");
			}

		}


		private void CheckAndAddToTable(ref Table ParentTable,ref Table checkTable,string altMsg,string OnClickMethod,string ImageSrc,string defaultStatus,string rowID)
		{
			if (checkTable.Rows.Count > 0)
			{
				PccRow myRow = new PccRow(); 
				myRow.SetDefaultCellData("",HorizontalAlign.Right,0,0);
				myRow.AddTextCell("<A title='" + altMsg + "'  onclick=" + OnClickMethod + " style='cursor:hand;'><img title='" + altMsg + "' src='" + Session["PageLayer"].ToString()  + "images/MenuArea/" + ImageSrc + "' border='' /></A>",80);
				myRow.AddTextCell(" ",20); 
				ParentTable.Controls.Add(myRow.Row);
  
				myRow.Reset();
				myRow.SetDefaultCellData("",HorizontalAlign.Right,0,2);
				myRow.SetRowID(rowID);
				myRow.SetRowCss(defaultStatus);
				myRow.AddControl(checkTable,100); 
				ParentTable.Controls.Add(myRow.Row);
			}
		}

		private string GetApFolder(string ApLink)
		{
			int pos = ApLink.IndexOf("/");
			string strReturn = ApLink.Substring(0,pos);
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
			j = Server.MapPath(".").IndexOf(Application["EDPNET"].ToString());

			try
			{
				for (i = 1 ; i < Server.MapPath(".").Substring(j).Split('\\').Length ; i++)
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

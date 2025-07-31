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
	///		Menus 的摘要描述。
	/// </summary>
	public partial  class Menus : System.Web.UI.UserControl
	{


		protected void Page_Load(object sender, System.EventArgs e)
		{
			// 將使用者程式碼置於此以初始化網頁
			if (Session["UserName"] != null && Session["UserName"].ToString()  != "")
			{
				//取得應用程式之APID
				string strApID = Request.Params["ApID"];
				string ApFolder = "";
				
				if (strApID == null)
					strApID = "";

				//TableRow Rows;
				TableCell Cells;

				//使得第一列與Menu上有一些距離
				PccRow myRow = new PccRow(); 
				myRow.AddTextCell(" ",0);
				tblSpace.Rows.Add(myRow.Row);  
				
				//取得此User所能使用之AP，動態產生主要的Menu
				int counts = 0;
				PccCommonForC.PccMsg myMsg = new PccCommonForC.PccMsg(); 

				tblMainMenu.CssClass = "topmenu";
				Session["SupAdmin"] = "N";

				myMsg.LoadXml(Session["XmlLoginInfo"].ToString());

				//判斷是否有系統可以使用，若沒有則直接跳出
				if (myMsg.QueryNodes("Authorize") != null)
				{ 
					myRow.Reset(); 
					//首先加入上層之Ap Menu，每超過五個就加入一列 2004/3/4
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

				//利用LoginInfo取得此User的AP Detail Menu.
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
            
				//使得第一列與圖形會有一定的距離
				myRow.Reset(); 
				Cells = new TableCell();
				Cells.Height = Unit.Pixel(20);
				myRow.Row.Cells.Add(Cells); 
				tblDetailMenu.Rows.Add(myRow.Row);
				
				//加入歡迎的前置詞，並讓使用者可以Link至UpdateLoginUser的網頁中來修改自己的資料
				myRow.Reset(); 
				myRow.SetDefaultCellData("CellWelcome",HorizontalAlign.Left,VerticalAlign.Middle,0);
				myRow.AddTextCell(myLabel.GetErrMsg("msgWelcome") + "<A href=" + strDetailPageLayer + "UpdateLoginUser.aspx>" + Session["UserName"] + "</A>" + myLabel.GetErrMsg("msgWelcome2"),0); 
				tblDetailMenu.Rows.Add(myRow.Row);
				
				if (strApID.Length > 0 && strDetailPageLayer.Length > 0)
				{
					strDetailPageLayer = strDetailPageLayer.Substring(0,strDetailPageLayer.Length - 3);

					//判斷是否有系統可以使用，若沒有則直接跳出
					if (myMsg.QueryNodes("Authorize") != null)
					{
						foreach (XmlNode myNode in myMsg.QueryNodes("Authorize"))
						{
							//假如不是點選這個系統則直接跳出並判斷下一個
							if (strApID != myMsg.Query("APID",myNode)) continue; 

							//取得這個AP的最上層之Folder
							ApFolder = GetApFolder(myMsg.Query("APLink",myNode)); 

							//新增這個系統的前置詞
							myRow.Reset();
							myRow.SetDefaultCellData("menuheading",0,0,2);
							myRow.AddTextCell(myMsg.Query("APName",(XmlElement)myNode) + myLabel.GetErrMsg("msgMenu"),0); 
							tblDetailMenu.Controls.Add(myRow.Row);

							//判斷是否有系統可以使用，若沒有則直接跳出
							if (myMsg.QueryNodes("ApMenu",(XmlElement)myNode) != null)
							{
								foreach(XmlNode myDetailNode in myMsg.QueryNodes("ApMenu",(XmlElement)myNode))
								{
									if (myMsg.Query("show_mk",myDetailNode) == "Y") 
									{
										myRow.Reset();
										myRow.SetDefaultCellData("menuitem",0,0,0);
										//判斷若是管理區則直接連接至SysManager的Floder中
										if (myMsg.Query("MenuManage",(XmlElement)myDetailNode) == "Y")
										{
											myRow.AddTextCell("<A href=" + "../" + strDetailPageLayer + "SysManager/" + myMsg.Query("MenuLink",(XmlElement)myDetailNode) + " >" + myMsg.Query("MenuNM",(XmlElement)myDetailNode) + "</A>",90);
										}
										else
										{
											myRow.AddTextCell("<A href=" + "../" + strDetailPageLayer + ApFolder + "/" + myMsg.Query("MenuLink",(XmlElement)myDetailNode) + " >" + myMsg.Query("MenuNM",(XmlElement)myDetailNode) + "</A>",90);
										}
										myRow.AddTextCell(" ",10); 
										//判斷Menu的分區，並加入分隔符號 20040108
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
			// CODEGEN: 此呼叫為 ASP.NET Web Form 設計工具的必要項。
			//

			//計算這個網頁所在的層次是那裡
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
		/// 此為設計工具支援所必需的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{
		}
		#endregion
		
	}
}

<%@ Page language="c#" Inherits="PLSignWeb2.Pub.Module.LoginTop1" CodeFile="LoginTop1.aspx.cs" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD runat="server">
		<title></title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="JavaScript" src="../js/common.js"></script>
		<script language="javascript" src="../js/DOL_CORE.js"></script>
		<script language="javascript" src="../js/DOL_MenuBar.js"></script>
		<script language="javascript" src="../js/DOL_XpItem.js"></script> 
		<link rel="stylesheet" href="../../Pub/Css/BodyStyles.css">
		<script language="javascript">
			function fun1()
			{
				var apID = GetApLink(document.all('txtTopMenu').value,"APName","APID",GetApName());
				top.frames[1].location.href("Medium1.aspx?ApID=" + apID);
			}
			function funReLogin(){        			                                             			                                             			                                                     
			   	chk = confirm("確認登出現有使用者？");                    
			      if(chk == true){window.open("../../Default.aspx","_top");}    
			      //if(chk == true){window.close();}        
			}
			  
			function funExit(){				
				top.window.close();
			}  
			
			function funPersionPage()
			{
				chk = confirm("確認回到個人化首頁？");                    
			      if(chk == true)
			      {
					top.location.href("http://172.16.1.25/personal_page/user_chk.asp?user=personal&upd_id=<%=Session["EncodeUpdID"]%>");
				  }
			}
			
			function pccLoad()
			{
				
				TsXpMenu.LogoWidth = 200;
					
				var arrApName = SecondLayer(document.all('txtTopMenu').value,"APName"); 
				var arrApID = SecondLayer(document.all('txtTopMenu').value,"APID");
				var oItem = new Array();
				
				var oMenu = new TsXpMenu("100%"); 
				//設定ApName Title
				var _logostr = "";
				_logostr += "<table cellspacing =0 cellpadding=0 border=0 width=100%><tr>";
				_logostr += "<td align=center class='dataTit01'>請決定系統名稱</td></td></table>";
				oMenu.setLogoHTML(_logostr); 
				
				var oBody = oMenu.getBody(); 
				
				oMenu.setHeight(60);
				
				for (i = 0 ; i < arrApName.length ; i++)
				{
					//因為系統不知道ApID為多少，所以先以系統管理的圖示來表示，若建立系統後，需把以下的Mark，並建立ApID的圖示，並使用變動之圖示即把下二行的Mark拿掉
					oItem[i] = new TsXpMenuItem(SplitApName(arrApName[i]),"../../Images/MenuArea/104.gif");
					//oItem[i] = new TsXpMenuItem(SplitApName(arrApName[i]),"../../Images/MenuArea/" + arrApID[i] +".gif");
					oItem[i].AddSelectEvent(fun1);
					oMenu.AddItem(oItem[i]);
				}
				
				var persionPageItem = new TsXpMenuItem("個人<br>首頁","../../Images/MenuArea/PersionalPage.gif");
				persionPageItem.AddSelectEvent(funPersionPage);
				oMenu.AddItem(persionPageItem);
				
				var reLoginItem = new TsXpMenuItem("重新<br>登錄","../../Images/MenuArea/menu_login.gif");
				reLoginItem.AddSelectEvent(funReLogin);
				oMenu.AddItem(reLoginItem);
				
				var ExitItem = new TsXpMenuItem("離開<br>系統","../../Images/MenuArea/menu_exit.gif");
				ExitItem.AddSelectEvent(funExit);
				oMenu.AddItem(ExitItem);
				
				document.body.appendChild(oBody);
			}
			
			function SplitApName(apName)
			{
				if ((apName.length < 3) || (apName.length > 4)) return apName;
				
				var strbegin = apName.substr(0,2);
				var strend = apName.substr(2);
				
				return strbegin + "<br>" + strend;
				
			}
			
			function GetApName()
			{
				//判斷是否是只有圖檔或是圖檔加文字
				var apName = event.srcElement.innerText;
				if (apName.length == 0)
					apName = event.srcElement.parentNode.innerText;
				
				var strtemp = "";
				
				for(i=0;i<apName.length;i++)
				{
					if (apName.charCodeAt(i) > 32)
					{
						strtemp = strtemp + apName.charAt(i);
					}
				}
				
				apName = strtemp;
				
				return apName;
				
			}
		
		</script>
	</HEAD>
	<body onload="pccLoad();" style="BACKGROUND:#f0f0f0;MARGIN:0px">
		<form method="post" runat="server">
			<input type="hidden" id="txtTopMenu" name="txtTopMenu" runat="server">
		</form>
	</body>
</HTML>

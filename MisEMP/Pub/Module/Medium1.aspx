<%@ Page language="c#" Inherits="PLSignWeb2.Pub.Module.Medium1" CodeFile="Medium1.aspx.cs" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD runat="server">
		<title></title>
		<meta content="Microsoft Visual Studio 7.0" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="JavaScript" src="../js/common.js"></script>
		<script language="javascript" src="../js/DOL_CORE.js"></script>
		<script language="javascript" src="../js/DOL_XpToolBar.js"></script>
		<script language="javascript" src="../js/DOL_XpTabBar.js"></script>
		<script language="javascript" src="../js/DOL_XpItem.js"></script>
		<script language="javascript" src="../js/DOL_XpProcessBar.js"></script>
		<script language="javascript" src="../js/DOL_DgyyWebWinNew.js"></script>
		<script language="javascript">
			function pccLoad()
			{
				var menuArray = document.all('txtLeftMenu').value;																
					
				if (menuArray != "")
				{
					eval(menuArray);	
					
					//在new DgyyWebWinNew之前設定工具欄之間的間隔距離
					DgyyWebWinNew.ToolBarInterval =1;
					
					//實例化界面對象(字符串數組,寬度,高度)
					oWin = new DgyyWebWinNew(menus,"100%","100%");
					
					for (i = 0; i < menus.length; i++)
					{
						//表示只有一般區會被打開
						if (i != 1)
							oWin.HideToolBar(i);
					}
					//顯示在body中
					oWin.Show();  
					
					oWin.OpenPage("訊息頁","LoginBody.aspx?ApID=0");
				}
				else
				{
					window.parent.location.href = '../../Default.aspx';
				}
			}
		</script>
	</HEAD>
	<body onload="pccLoad();" scroll=no style="BACKGROUND:#f0f0f0;MARGIN:0px">
		<form method="post" runat="server">
			<input type="hidden" id="txtLeftMenu" name="txtLeftMenu" runat="server">
		</form>
	</body>
</HTML>

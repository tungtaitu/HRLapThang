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
					
					//�bnew DgyyWebWinNew���e�]�w�u���椧�������j�Z��
					DgyyWebWinNew.ToolBarInterval =1;
					
					//��ҤƬɭ���H(�r�Ŧ�Ʋ�,�e��,����)
					oWin = new DgyyWebWinNew(menus,"100%","100%");
					
					for (i = 0; i < menus.length; i++)
					{
						//��ܥu���@��Ϸ|�Q���}
						if (i != 1)
							oWin.HideToolBar(i);
					}
					//��ܦbbody��
					oWin.Show();  
					
					oWin.OpenPage("�T����","LoginBody.aspx?ApID=0");
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

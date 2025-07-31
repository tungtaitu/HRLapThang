<%@ Page language="c#" Inherits="PLSignWeb2.Bottom" CodeFile="LoginBottom.aspx.cs" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD runat="server">
		<title></title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script language="JavaScript" src="../js/common.js"></script>
		<link rel="stylesheet" href="../css/PccStyles.css">
		<script language="javascript">  
		function Hello()
			{
				window.open("<%=System.Configuration.ConfigurationSettings.AppSettings["myServer"]%>" + "<%=System.Configuration.ConfigurationSettings.AppSettings["vpath"]%>" + "/OnlineUser.aspx?Type=Logout&Type2=Close","new","width=600,height=600,top=100,left=100,toolbar=no,location=no,status=no"); 
				
			}
		</script>
	</HEAD>
	<body>
		<form method="post" runat="server">
			<FONT face="新細明體">線上人數：<%=Application["OnlineCount"]%>&nbsp;&nbsp;
			<% if (KeyValue["superAdminEmail"].Trim().ToLower() == Session["UserEMail"].ToString().Trim().ToLower())
      { %>
				<input type=button class=button value="OnlineUser" onclick="Hello()" /> 
			<% } %> 
			</FONT>&nbsp; 
			<A id="lnkContact" href="mailto:<%=KeyValue["System-Emai"]%>?subject=系統問題反應" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','../../Images/email1.gif',1)" class="A1">
			<img src="../../Images/email2.gif" border="0" name="Image4" alt="連絡我們">有問題，請反應，Thanks！</A>
		</form>
	</body>
</HTML>

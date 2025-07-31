<%@ Control Language="c#" Inherits="PLSignWeb2.Pub.Module.Menus" CodeFile="Menus.ascx.cs" %>
<%@ OutputCache Duration="2" VaryByParam="None" VaryByCustom="authenticated" %>
<mata HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=big5">
	<script SRC="<%=Session["PageLayer"]%>Pub/js/common.js" LANGUAGE="JavaScript">
	</script>
	<link rel="stylesheet" href="<%=Session["PageLayer"]%>Pub/css/PccStyles.css">
	<input type=hidden id=txtAuthXml value=''>
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td bgColor="#78a0be"><IMG height="1" src="<%=Session["PageLayer"]%>images/space.gif" width="155" border="0"></td>
		<td bgColor="#78a0be"><IMG height="1" src="<%=Session["PageLayer"]%>images/space.gif" width="128" border="0"></td>
		<td bgColor="#78a0be"><IMG height="1" src="<%=Session["PageLayer"]%>images/space.gif" width="500" border="0"></td>
		<td bgColor="#78a0be"><IMG height="1" src="<%=Session["PageLayer"]%>images/space.gif" width="*" border="0"></td>
	</tr>
	<TR>
		<TD class="topLeft1" height="65"></TD>
		<TD class="topLeft2" height="65"></TD>
		<TD class="topRight" valign="bottom" height="65">
			<asp:Table id="tblMainMenu" runat="server" Width="450px" Height="8px"></asp:Table>
		</TD>
		<TD class="topRight" vAlign="bottom" height="65">
			<TABLE class="topmenu">
				<TR>
					<TD><A id="lnkHome" href="<%=Session["PageLayer"]%>PccApHome.aspx">Home </A>
					</TD>
					<TD><A id="lnkLogout" href="<%=Session["PageLayer"]%>PccLogin.aspx">Logout </A>
					</TD>
					<TD><A id="lnkAbout" href="<%=Session["PageLayer"]%>PccAbout.aspx">About </A>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<td colspan="1" background="<%=Session["PageLayer"]%>images/pcc_7.jpg" width="155" height="100%" valign="top">
			<table background="<%=Session["PageLayer"]%>images/pcc_5.jpg" width="155" height="394" border="0" cellpadding="0" cellspacing="0" style="BACKGROUND-ATTACHMENT: fixed; BACKGROUND-REPEAT: no-repeat">
				<tr>
					<td valign="top">
						<asp:Table ID="tblDetailMenu" Runat=server Width="100%"></asp:Table> 
					</td>
				</tr>
			</table>
		</td>
		<td rowspan="3" colspan="3" width="*" valign="top" background="<%=Session["PageLayer"]%>images/bg-earth.gif">
			<asp:Table id="tblSpace" runat="server" Height="8px"></asp:Table>

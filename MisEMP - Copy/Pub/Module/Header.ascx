<%@ Control Language="c#" Inherits="PLSignWeb2.Pub.Module.Header" CodeFile="Header.ascx.cs" %>
	<script SRC="<%=Session["PageLayer"]%>Pub/js/common.js" LANGUAGE="JavaScript">
	</script>
	<link rel="stylesheet" href="<%=Session["PageLayer"]%>Pub/css/BodyStyles.css">
	<script language="javascript">
			window.onload = window_common_onload;

			function window_common_onload()
			{
				if ( top.document.all )
				{
					//alert('<%=Session["PageLayer"]%>');
					if ('<%=Session["UserID"]%>' == "")
					{ 
						window.parent.parent.location.href = '<%=Session["PageLayer"]%>' + 'Default.aspx';  
					}
				}
			}
	
	</script>
	<table id="table1" cellpadding = "0" cellspacing = "0" border = 0 width = "100%"  background="../../images/bg_title.gif">
	<tr>
		<td>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<asp:label id="PccTitle" runat="server" Width="100%" Font-Bold="True" ForeColor ="white"  Font-Size="10pt">Title</asp:label>
		</td>
	</tr>
	</table>
	<table id="table2" cellpadding = "0" cellspacing = "0" border = 0 width = "100%" height ="5px">
	<tr>
		<td></td>
	</tr>
	</table>
	
	

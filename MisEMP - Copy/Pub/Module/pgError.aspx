<%@ Page language="c#" Inherits="PLSignWeb2.Pub.Module.pgError" CodeFile="pgError.aspx.cs" %>
<%@ Register TagPrefix="uc1" TagName="Menus" Src="Menus.ascx" %>
<%@ Register TagPrefix="uc1" TagName="Menus_End" Src="Menus_End.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD runat="server">
		<title></title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio 7.0">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body>
		<form method="post" runat="server">
			<FONT face="新細明體">
				<P align="left">
					<TABLE id="Table1" style="WIDTH: 611px; HEIGHT: 145px" borderColor="#cc0033" cellSpacing="0" cellPadding="0" width="611" border="1">
						<TBODY>
							<TR>
								<TD class="ActDocTD3">
									<P><STRONG><FONT style="FONT-SIZE: 14pt" color="background">系統拒絕登錄</FONT></STRONG></P>
									<P style="FONT-SIZE: 10pt"><FONT size="1"></P>
									<P style="FONT-SIZE: 10pt">
			</FONT>可能原因為登錄資料錯誤、資料尚未建檔、或該使用者無使用此系統之權限。</P>
			<P style="FONT-SIZE: 10pt"><FONT size="1"><FONT face="新細明體"></FONT>&nbsp;</P>
			</FONT>
			<P align="center">
				．
				<asp:HyperLink id="reLogin" runat="server" ForeColor="Blue" Font-Size="10pt">重新登錄</asp:HyperLink>．</P>
			</TD></TR></TBODY></TABLE></P></FONT>
		</form>
	</body>
</HTML>

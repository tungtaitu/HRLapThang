<%@ Page language="c#" Inherits="PLSignWeb2.Pub.Module.UserLogin" CodeFile="UserLogin.aspx.cs" %>
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
			<FONT face="�s�ө���">
				<TABLE id="Table1" style="WIDTH: 361px; HEIGHT: 293px" cellSpacing="0" cellPadding="0" width="361" align="center" border="0">
					<TR>
						<TD><FONT face="�s�ө���"><A href="default.asp"><IMG src="../../Images/pcc_logo.gif" border="0"></A></FONT></TD>
					</TR>
					<TR>
						<TD>
							<P align="center"><FONT size="2">�w��Ө��_��������T�t�ΡA�ϥΪ̵n�J����<BR>
									���F�z�ϥΪ��v�q�A�бz�n���n�J�t�γ�!</FONT></P>
						</TD>
					</TR>
					<TR>
						<TD><FONT face="�s�ө���"></FONT></TD>
					</TR>
					<TR>
						<TD>
							<DIV align="center">
								<TABLE id="Table2" style="WIDTH: 324px; HEIGHT: 121px" cellSpacing="0" cellPadding="0" width="324" align="center" background="../../Images/login_bg.gif" border="0">
									<TR>
										<TD style="WIDTH: 101px; HEIGHT: 34px" colSpan="1" rowSpan="1"><FONT face="�s�ө���"></FONT></TD>
										<TD style="WIDTH: 44px; HEIGHT: 34px"><FONT face="�s�ө���"></FONT></TD>
										<TD style="WIDTH: 169px; HEIGHT: 34px"></TD>
									</TR>
									<TR>
										<TD style="WIDTH: 101px; HEIGHT: 18px"><FONT face="�s�ө���"></FONT></TD>
										<TD style="WIDTH: 44px; HEIGHT: 18px" width="44" colSpan="1" rowSpan="1"><FONT face="�s�ө���" size="2">�b���G</FONT></TD>
										<TD style="WIDTH: 169px; HEIGHT: 18px"><FONT face="�s�ө���" size="2">
												<asp:TextBox id="txtUserName" runat="server" Height="15pt" BorderWidth="1px" BorderColor="RoyalBlue" BorderStyle="Solid"></asp:TextBox></FONT></TD>
									</TR>
									<TR>
										<TD style="WIDTH: 101px; HEIGHT: 13px"><FONT size="2"></FONT></TD>
										<TD style="WIDTH: 44px; HEIGHT: 13px"><FONT face="�s�ө���" size="2">�K�X�G</FONT></TD>
										<TD style="WIDTH: 169px; HEIGHT: 14px"><FONT face="�s�ө���">
												<asp:TextBox id="txtPassWord" runat="server" Height="15pt" BorderWidth="1px" BorderColor="RoyalBlue" TextMode="Password"></asp:TextBox></FONT></TD>
									</TR>
									<TR>
										<TD style="WIDTH: 315px; HEIGHT: 35px" colSpan="3">
											<P>
												<TABLE id="Table3" style="WIDTH: 322px; HEIGHT: 29px" cellSpacing="0" cellPadding="0" width="322" align="center" border="0">
													<TR>
														<TD style="WIDTH: 300px; HEIGHT: 29px" align="middle">
															<P>
																<asp:Button id="btnLogin" runat="server" Font-Size="8pt" Text="�n�J" onclick="btnLogin_Click"></asp:Button>&nbsp;
																<asp:Button id="btnClear" runat="server" Font-Size="8pt" Text="�M��" onclick="btnClear_Click"></asp:Button></P>
														</TD>
													</TR>
												</TABLE>
											</P>
											<P align="center">
												<asp:Label id="lblErrorMsg" runat="server" Width="319px" Font-Size="10pt" Font-Bold="True" ForeColor="Maroon"></asp:Label></P>
										</TD>
									</TR>
								</TABLE>
							</DIV>
						</TD>
					</TR>
				</TABLE>
			</FONT>
		</form>
	</body>
</HTML>

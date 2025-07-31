<%@ Language=VBScript %>
<%
text1=Request.Form("text1").Item
text2Form=Request.Form("text2").Item
text2URL=Request.QueryString("text2").Item
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function button1_onclick() {
	var text2=window.document.form1.text2.value;
	var strURL='test.asp?text2=' + text2;
	window.document.form1.action=strURL;
	window.document.form1.submit();
}

//-->
</SCRIPT>
</HEAD>
<BODY>
<FORM action="test.asp" method=POST id=form1 name=form1>
text1:<%=text1%><br>
text2Form:<%=text2Form%><br>
text2URL:<%=text2URL%><br>
<hr>

<INPUT type="text" id=text1 name=text1 value="¯]Ä_"><br>
<INPUT type="text" id=text2 name=text2 value="<%=Server.URLEncode("¯]Ä_")%>"><br>
<INPUT type="button" value="Button" id=button1 name=button1 LANGUAGE=javascript onclick="return button1_onclick()">
</FORM>
</BODY>
</HTML>

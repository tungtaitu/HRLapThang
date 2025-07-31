<%@Language=VBScript codepage=65001 %>
<!--#include file="../include/sideinfo.inc"-->

<% 
Response.Expires = 0 
Response.Buffer=true
FUNC = Trim(Request("func"))
%>
<HTML>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	


</head>
<BODY   topmargin=0>

<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

<table width=640 border=0 ><tr><td align=center>
<FORM method=post action="<%=Session("Action")%>" name="<%=Session("Name")%>" target="_self" >
<INPUT type=submit name=submit value="<%=replace(Session("SubmitValue"),"<BR>",chr(13))%>" class="btn btn-sm btn-danger">
</FORM>
</td>
</tr>
<tr><td align=center>
<FONT COLOR=blue class="txt12"><%=Session("NO")%></FONT><br>
<FONT COLOR=blue class="txt12"><%=Session("KeyValue")%></FONT><br>
<FONT COLOR=blue class="txt12"><%=Session("MessageCode")%></FONT>
</td></tr></table>

<%
' FormName=Session("Name")
' Session("Title") = ""
' Session("Name") = ""
' Session("MessageCode") = ""
' Session("KeyValue") = ""
' Session("Target") = ""
' Session("SubmitValue") = ""
' Session("Action") = ""
%>


</BODY>
</HTML>

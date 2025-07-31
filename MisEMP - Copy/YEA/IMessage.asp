<%@language=vbscript codepage=65001%>
<% 
Response.Expires = 0 
Response.Buffer=true
FUNC = Trim(Request("func"))
%>
<script language="vbscript">
	function global()
	end function
</script>

<HTML>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">	
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<BODY onload="global()" background="bg_blue.gif">
<table width=500><tr><td>
<center>
  <font size="5"><b><%=Session("Title")%></b></font> 
</center>
<hr><br><br><br><br>
<center>  
<FORM method=post action="<%=Session("Action")%>" name="<%=Session("Name")%>" >

<INPUT type=submit name=submit value="<%=Session("SubmitValue")%>">
</FORM>

<P>
<center><FONT COLOR=blue class="txt12"><%=Session("NO")%></FONT></center>
<center><FONT COLOR=blue class="txt12"><%=Session("KeyValue")%></FONT></center>
<center><FONT COLOR=blue class="txt12"><%=Session("MessageCode")%></FONT></center>
</center>
</td></tr></table>
<%
FormName=Session("Name")
Session("Title") = ""
Session("Name") = ""
Session("MessageCode") = ""
Session("KeyValue") = ""
Session("Target") = ""
Session("SubmitValue") = ""
Session("Action") = ""
%>
</BODY>
</HTML>

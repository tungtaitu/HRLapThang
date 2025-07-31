<%@LANGUAGE=VBSCRIPT codepage=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
SELF="YECE0801" 
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "<%=SELF%>.FORE.asp"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			noresize>
	<frame	SRC = "getEmpData.asp?formName=<%=self%>" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




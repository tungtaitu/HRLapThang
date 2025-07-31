<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
SELF="YEGBE0102"
%>
<HTML>
<HEAD>
<title></title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id=best>  
	<frame	SRC = "<%=SELF%>.Fore.asp"  
			name="Main"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			noresize>
	<frame	SRC = "<%=SELF%>.Back.asp" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>

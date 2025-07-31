<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
self="YEGEE0401" 

%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="50%,50%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "<%=self%>.Fore.asp"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			>
	<frame	SRC = "<%=self%>.back.asp" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




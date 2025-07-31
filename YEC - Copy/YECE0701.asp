<%@language=vbscript codepage=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"--> 
<%
self="YECE0701"
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "<%=self%>.FORE.asp?pgid=<%=request("pgid")%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			noresize>
	<frame	SRC = "<%=self%>.Back2.asp" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




<%@language=vbscript codepage=65001%>
<%Response.Buffer =True%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"--> 
<%
SELF="YEAAE0301"
%>
<HTML>
<HEAD>
<title></title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="BEST">  
	<frame	SRC = "<%=SELF%>.FORE.ASP?pgid=<%=request("pgid")%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling=auto 
			noresize>
	<frame	SRC = "" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling=auto>	
  </frameset>
</HTML>

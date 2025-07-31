<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!--#include file="../../GetSQLServerConnection.fun"-->
<!--#include file="../../include/checkpower.asp"-->  
 
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "empbasic.Fore.asp?pgid=<%=request("pgid")%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			noresize>
	<frame	SRC = "empbasic.Back.asp?pgid=<%=request("pgid")%>" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




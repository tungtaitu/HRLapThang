<%@language=vbscript codepage=65001%>
<%Response.Buffer =True%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
'response.write  session("mode") 
'response.end  
%>
<HTML>
<HEAD>
<title></title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="BEST">  
	<frame	SRC = "YSBAE0401.FORE.ASP"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling=no 
			noresize>
	<frame	SRC = "" 
			name = "help"  
			framespacing=0 
			frameborder=0 
			scrolling=auto>	
  </frameset>
</HTML>

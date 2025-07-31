<%@language=vbscript codepage=65001%>
<%Response.Buffer =True%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
'response.write  session("mode") 
'response.end  
SELF="YEAAE0401"
%>
<HTML>
<HEAD>
<title></title>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="BEST">  
	<frame	SRC = "<%=SELF%>.FORE.ASP?pgid=<%=request("pgid")%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling=auto 
			noresize>
	<frame	SRC = "<%=SELF%>.Back.ASP" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling=auto>	
  </frameset>
</HTML>

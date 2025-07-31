<%@language=vbscript codepage=65001%>

<%
self="YTBAE01"
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%, 0%" frameborder = "NO" framespacing=0 id="BEST">  
	<frame	SRC = "<%=self%>.Forgnd.asp?pgid=<%=request("pgid")%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling=auto 
			noresize>
	<frame	SRC = "<%=self%>.Back.asp"
			name="Back"  
			framespacing=0 
			frameborder=0 
			scrolling=auto 
			noresize>		
  </frameset>

</HTML>

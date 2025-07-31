<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<%self="empholiday"
cf=request("cf")
empid=request("empid")
%> 
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "<%=self%>.fore.asp?cf=<%=cf%>&empid=<%=empid%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			>
	<frame	SRC = "<%=self%>.Back.asp" 
			name="Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




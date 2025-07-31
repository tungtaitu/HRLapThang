<%@LANGUAGE=VBSCRIPT CODEPAGE=950%>
<%self="empfiledt" 
'response.write "aaa"
empautoid = request("empautoid") 
totalpage = request("totalpage")  
currentpage = request("currentpage")
RecordInDB = request("RecordInDB")

%> 
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=BIG5">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="100%,0%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC="empfile.foregnd.asp?empautoid=<%=empautoid%>&totalpage=<%=totalpage%>&currentpage=<%=currentpage%>&RecordInDB=<%=RecordInDB%>"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			noresize>
	<frame	SRC = "<%=self%>.Back.asp" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!--#include file="../GetSQLServerConnection.fun"-->
<!--#include file="../include/checkpower.asp"-->  
<%
self="foewait" 

sdate = request("DAT1")
edate = request("DAT2")
eid = request("eid") 
'response.write sdate & edate & eid 
%>
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
</HEAD>
  <frameset cols="50%,50%" frameborder = "NO" framespacing=0 id="best">  
	<frame	SRC = "forwait_tmp2.asp"
			name="Fore"  
			framespacing=0 
			frameborder=0 
			scrolling="auto" 
			>
	<frame	SRC = "YEGEE0301.foregnd.asp?dat1=<%=sdate%>&dat2=<%=edate%>&eid=<%=eid%>" 
			name = "Back"  
			framespacing=0 
			frameborder=0 
			scrolling="auto">
  </frameset>

</HTML>




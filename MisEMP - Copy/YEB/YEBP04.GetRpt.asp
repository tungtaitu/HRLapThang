<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
const self ="yebp04.getrpt.asp"
%>
<%
reportname="empcard.rpt"
code01=request("whsno")
code02=request("groupid")
code03=request("country")
code04=request("JOB")
code05=request("empid1")
code06=request("empid2")
code07=request("indat1")
code08=request("indat2")
code09=request("empTJ")
code10=request("bhdat1")
code11=request("bhdat2") 
inym = request("inym")
loai=request("loai") 

if code02="WB" then 
	response.redirect "YEBP04WB.getrpt.asp?whsno="& code01 &"&loai="& loai &"&empid1="& code05 & "&empid2="& code06
else	
	reportname="empcard.rpt"
end if 	



%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code06))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(inym))
%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->






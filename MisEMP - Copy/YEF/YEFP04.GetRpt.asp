<%
const self ="emp_worktime_bydate.getrpt.asp"
%>
<%
'reportname = "emp_worktime_bydate.rpt"
viewid = ucase(session("netuser")) 
'viewid="LSARY" 

if viewid="LSARY" THEN 
	reportname = "emp_workTime_Bydate_JD.rpt"
else
	reportname = "emp_worktime_bydateN.rpt"
end if 	

code01=request("whsno")
code02=request("groupid")
code03=request("country")
code04=request("JOB")
code05=request("empid1")
code06=request("empid2")
code07=replace(request("indat1"),"/","")
code08=replace(request("indat2"),"/","")
code09=request("outemp")
code10=request("shift")
code11=request("showby")
zuno=replace(trim(request("zuno"))," ","")

response.write code01&"<Br>"
response.write code02&"<Br>"
response.write code03&"<Br>"
response.write code04&"<Br>"
response.write code05&"<Br>"
response.write code06&"<Br>"
response.write code07&"<Br>"
response.write code08&"<Br>"
response.write code09&"<Br>"
response.write code10&"<Br>"
response.write code11&"<Br>"
response.write zuno&"<Br>"

Response.write  reportname
response.end

on error resume next 		  

%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" --> 
<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code08))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code10))
Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code11))
Session("oRpt").ParameterFields(10).AddCurrentValue(Cstr(zuno))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








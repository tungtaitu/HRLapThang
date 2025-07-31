<%
'const self ="emp_BreakORKwime.getrpt.asp"
%>
<%
'reportname = "emp_BreakTime.rpt"
RESPONSE.WRITE reportname

code01=request("whsno")
code02=request("groupid")
code03=request("country")
code04=request("JOB")
code05=request("empid1")
code06=request("empid2")
code07=replace(request("indat1"),"/","")
code08=replace(request("indat2"),"/","")
code09=request("outemp")

'response.write code07&"<RB>"
'response.write code08

'response.end

sy  = request("sy")
if sy = "A" then 
	reportname = "emp_BreakTime_empidN.rpt"
else
	reportname = "emp_BreakTime_groupid.rpt"
end if 
%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code08))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code05))
'Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code06))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








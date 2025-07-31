<%
const self ="emp_worktimeN.getrpt.asp"
%>
<%
reportname = "emp_workTimeToT_N.rpt"

code01= trim(replace(request("D1"),"/",""))
code02= Trim(replace(request("D2"),"/",""))
code03=replace(request("country")," ","")
code04=request("whsno")
code05=replace(request("groupid")," ","")
code06=request("empid1")
code07=request("zuno")
code08=request("shift")
userid=session("netuser")
'response.write code01&"<BR>"
'response.write code02&"<BR>"
'response.write code03&"<BR>"
'response.write code04&"<BR>"
'response.write code05&"<BR>"
'response.write code08 
'response.write userid 
'response.end 

%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(userid))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








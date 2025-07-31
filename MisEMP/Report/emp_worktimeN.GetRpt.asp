<%
const self ="emp_worktimeN.getrpt.asp"
%>
<%
reportname = "emp_workTimeToT.rpt"

code01= trim(replace(request("D1"),"/",""))
code02= Trim(replace(request("D2"),"/",""))
code03=request("country")
code04=request("whsno")
code05=request("groupid")
code06=request("empid1")

'response.write code01&"<BR>"
'response.write code02&"<BR>"
'response.write code03&"<BR>"
'response.write code04&"<BR>"
'response.write code05&"<BR>"
'response.write code08 
'response.write reportname 
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
'Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








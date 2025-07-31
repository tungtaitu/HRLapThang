<%
const self ="salarycp12.getrpt.asp"
%>
<%
reportname = "empsalaryCP12.rpt"


code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("empid1")
code06=request("SY")


'response.write code01&"<RB>"
'response.write code02&"<RB>"
'response.write code03&"<RB>"
'response.write code04&"<RB>"
'response.write code05&"<RB>"
'response.write code06&"<RB>" 

'response.end


%>

<!-- #include file="../../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code05))
'Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code04))
'Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code05))
'Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
'Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








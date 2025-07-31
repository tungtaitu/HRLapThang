<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<%
const self ="emp_holiday.getrpt.asp"
%>
<%
reportname = "emp_holiday.rpt"


code01=request("dat1")
code02=request("dat2")
code03=request("country")
code04=request("whsno")
code05=request("groupid")
code06=request("job")
code07=request("empid1")
 

'response.write code07&"<RB>"
'response.write code08 

'response.end 


%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








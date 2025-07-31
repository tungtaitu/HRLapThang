<%
const self ="salarycp04.getrpt.asp"
%>
<%
reportname = "EmpMsalary.rpt"


code01=request("YYMM")
code02=trim(request("country"))
code03=trim(request("whsno"))
code04=request("groupid")
code05=request("job")
code06=request("empid1")
 

'response.write code01&"<RB>"
'response.write code02&"<RB>"
'response.write code03&"<RB>"
'response.write code04&"<RB>"
'response.write code05&"<RB>"
'response.write code06&"<RB>"
'response.write code08 

'response.end 


%>

<!-- #include file="../../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code08))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code09))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








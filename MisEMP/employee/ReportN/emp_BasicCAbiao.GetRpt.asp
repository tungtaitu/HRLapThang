<%
const self		="emp_baiscCAbaio.getrpt.asp"
%>
<%
reportname = "emp_BasicCAbiao.rpt"


code1=request("empid")
%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code1))
'Session("oRpt").ParameterFields(2).AddCurrentValue(cstr(iym))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








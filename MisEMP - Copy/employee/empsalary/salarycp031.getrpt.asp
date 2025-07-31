<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
<!--#include virtual="yfy/ysb/include/ADOINC.inc"-->
<%
const self ="salarycp031.getrpt.asp"
%>
<% 

Set conn = GetSQLServerConnection()   

code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("empid1") 

'if code02="TA"  then 
'	reportname = "EMPDsalaryTA_N.rpt"
'elseif  code02="CN" then  
'	reportname = "EMPDsalaryTA.rpt"
'else
'	reportname = "EMPDsalary.rpt"
'end if	

reportname = "empNZJJ.rpt"
'response.write code07&"<RB>"
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
'Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
'Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








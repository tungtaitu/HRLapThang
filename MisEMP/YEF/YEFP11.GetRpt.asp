<%
const self ="emp_worktime.getrpt.asp"
%>
<%
'reportname = "emp_worktime.rpt" 

reportname = "emp_khb.rpt"
 

yymm=request("yymm")
code01=request("whsno")
code02=request("groupid")
code03=request("country")
code04=request("JOB")
code04=""
code05=request("empid1")
code06=request("empid2")
code07=replace(request("indat1"),"/","")
code08=replace(request("indat2"),"/","")
code09=request("outemp")
code10=request("shift")
code11=session("netuser")
zuno=request("zuno")
if request("yymm")="" then  yymm=left(code07,6)
'response.write code07&"<RB>"
'response.write zuno

'response.end


%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code08))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(zuno))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code10))
 

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








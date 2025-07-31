<%
const self ="salarycp02.getrpt.asp"
%>
<%

code01=request("YYMM")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1")
code07=request("outemp")
acc=request("acc")
nojx=request("nojx") 

'response.write  code05
'response.write  code06


if request("country")="VN"  then
	reportname = "prtempsalary.rpt"
else
	reportname = "prtempsalaryCN.rpt"
end if


'response.write code07&"<RB>"
'response.write code08



uid = session("netuser")

'response.write reportname
'response.end 
%>

<!-- #include file="../../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(uid))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(nojx))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code06)) 
Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code07)) 
Session("oRpt").ParameterFields(10).AddCurrentValue(Cstr(acc))


%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








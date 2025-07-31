<%
const self ="salarycp08.getrpt.asp"
%>
<%


code01=request("YYMM")
code02=request("JXYM")
code03=request("country")
code08=request("whsno")
code04=request("groupid")
code05=request("zuno")
code06=request("empid1")
code07=request("SHIFT")
sortby = request("sortby")
op=request("op")

if request("loai")="T" then 
	reportname = "../../REPORT/VYFYMYJX_TB.rpt" 
else
	reportname = "../../REPORT/VYFYMYJX.rpt" 
end if 	
'response.write code01&"<RB>"
'response.write code02&"<RB>"
'response.write code03&"<RB>"
'response.write code04&"<RB>"
'response.write code05&"<RB>"
'response.write code06&"<RB>"
'response.write code07&"<RB>"
'response.write code08

'response.end


%>

<!-- #include file="../../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(op))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code02))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code03))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code08))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code04))
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code05))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code06))
Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(10).AddCurrentValue(Cstr(sortby))

%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








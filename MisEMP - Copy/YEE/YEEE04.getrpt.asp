<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<%
Set conn = GetSQLServerConnection()

 
if  instr(conn,"168")>0 then 
	w1="LA"
	w2 = "越南"	
elseif  instr(conn,"169")>0 then 
	w1="DN"	
	w2 = "同奈"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
	w2 = "越南"	
end if 

set conn=nothing 

code01=request("xid") 
code02=request("empid1")
'code01=request("YYMM")
'code02=request("country")
'code03=request("whsno")
'code04=request("groupid")
'code05=request("job")
'code06=request("empid1") 
 
 
 
reportname = "PhieuCongTac.rpt"
 

'response.write code07&"<RB>"
'response.write code08 

'response.end 


%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(w1))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code02))
'Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(code03))
'Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code04))
'Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code05))
'Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
'Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code07))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








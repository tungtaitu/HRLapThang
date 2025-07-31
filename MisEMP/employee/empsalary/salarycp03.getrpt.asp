<!-------- #include file = "../../GetSQLServerConnection.fun" --------->
<!--#include virtual="yfy/ysb/include/ADOINC.inc"-->
<%
const self ="salarycp03.getrpt.asp"
%>
<% 

Set conn = GetSQLServerConnection()   

code01=request("YYMM")
acc=request("acc")
code02=request("country")
code03=request("whsno")
code04=request("groupid")
code05=request("job")
code06=request("empid1") 
code07=request("outemp") 
code08=request("acc") 

if code06<>"" then 
	sql="select * from view_empfile where empid='"& code06 &"' "
	set rds=conn.execute(sql) 
	if not rds.eof then 
		code02=rds("country")
	else
		code02=code02	
	end if 
else
	code02=code02 
end if 		 

if acc<>"" then 	
	reportname = "EMPDsalaryHW.rpt"	
else  
	if code02="TA"  then 
		reportname = "EMPDsalaryTA_N.rpt"
	elseif  code02="CN" then  
		reportname = "EMPDsalaryCN.rpt"
	elseif code02="VN" then 
		reportname = "EMPDsalary.rpt" 
	else
		response.end 	
	end if	 
end if 	

'response.write code03&"<RB>"
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
Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))
Session("oRpt").ParameterFields(8).AddCurrentValue(Cstr(code08))
'Session("oRpt").ParameterFields(9).AddCurrentValue(Cstr(code08))

%>
<!-- #include file="../../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../../global_report/ActiveXViewer.asp" -->








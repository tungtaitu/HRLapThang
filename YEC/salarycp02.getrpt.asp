<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
const self ="salarycp02.getrpt.asp"
%>
<%
Set conn = GetSQLServerConnection()	   
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
	reportname = "prtempsalary_VN.rpt"
else
	reportname = "prtempsalary_HW.rpt"
end if

nowmonth = year(date())&right("00"&month(date()),2)  

if code01<>nowmonth then 
	sql="select * from closeYM where closeym='"& code01 &"'"
	set rs=conn.execute(sql) 
	'response.write sql 
	'response.end 
	if rs.eof then 
		response.write "本月未關帳!! 請先 執行 C.B / 1.關帳(資料備份) 才可列印薪資單 !! "
		response.end 
	end if 
	set conn=nothing  
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








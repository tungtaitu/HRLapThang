<%

reportname = "emphd_vn.rpt"


code01=request("empid1")  
code02=left(request("indat1"),4) 
code03=mid(request("indat1"),6,2) 
code04=right(request("indat1"),2) 
code05=left(request("indat2"),4) 
code06=mid(request("indat2"),6,2)  
code07=right(request("indat2"),2) 
yymm=request("yymm")
whsno=request("whsno")
country=request("country")
date1=left(yymm,4)&"/"&right(yymm,2)&"/01"
cDat="" 
d1 = request("indat1")
d2 = request("indat2")

'response.write d1 & ","&  d2 & ","& whsno &  ","& country &  ","& code01 &  ","&  cDat 
%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<% 
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(d1))
Session("oRpt").ParameterFields(2).AddCurrentValue(Cstr(d2))
Session("oRpt").ParameterFields(3).AddCurrentValue(Cstr(whsno))
Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(country))
Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code01))
Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(cDat))
'Session("oRpt").ParameterFields(4).AddCurrentValue(Cstr(code01))
'Session("oRpt").ParameterFields(5).AddCurrentValue(Cstr(code05))
'Session("oRpt").ParameterFields(6).AddCurrentValue(Cstr(code06))
'Session("oRpt").ParameterFields(7).AddCurrentValue(Cstr(code07))

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








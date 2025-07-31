 
<% 

 
	reportname = "yebbp0201.rpt"
 


code01=request("yymm")
code02=request("groupid")
code03=request("country")
code04=request("JOB")
code05=request("empid1")
code06=request("empid2")
code07=request("indat1")
code08=request("indat2")
code09=request("empTJ")
code10=request("bhdat1")
code11=request("bhdat2")
code12=request("outemp")




%>

<!-- #include file="../global_report/AlwaysRequiredSteps.asp" -->
<!-- #include file="../global_report/OdbcConnection.asp" -->

<%
Session("oRpt").ParameterFields(1).AddCurrentValue(Cstr(code01))
 

%>
<!-- #include file="../global_report/MoreRequiredSteps.asp" -->
<%'--------需更改activexviewer.asp中的rptpath%>
<!-- #include file="../global_report/ActiveXViewer.asp" -->








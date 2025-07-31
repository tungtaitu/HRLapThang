<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<%
SELF = "YECE0901"

ftype = request("ftype")
code = request("code")
index=request("index")
CurrentPage = request("CurrentPage")

yymm=request("yymm")
 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
Cdays = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
response.write days

CODESTR01 = REQUEST("CODESTR01")
CODESTR02 = REQUEST("CODESTR02")
CODESTR03 = REQUEST("CODESTR03")
CODESTR04 = REQUEST("CODESTR04")
CODESTR05 = REQUEST("CODESTR05")
CODESTR06 = REQUEST("CODESTR06")
CODESTR07 = REQUEST("CODESTR07")
CODESTR08 = REQUEST("CODESTR08")
CODESTR09 = REQUEST("CODESTR09")
CODESTR10 = REQUEST("CODESTR10")
CODESTR11 = REQUEST("CODESTR11")
CODESTR12 = REQUEST("CODESTR12")
CODESTR13 = REQUEST("CODESTR13")
CODESTR14 = REQUEST("CODESTR14")
workdays = REQUEST("days")
response.write  "CODESTR13=" & CODESTR13 &"<BR>"

tmpRec = Session("YECE0901F")
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
 

Set conn = GetSQLServerConnection()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head>
<%
select case ftype  
    CASE "ZXCHG"
        tmpRec(CurrentPage,index + 1,0)  = "upd"
        tmpRec(CurrentPage,index + 1,23) = CODESTR02
        tmpRec(CurrentPage,index + 1,24) = CODESTR01   
        response.write tmpRec(CurrentPage,index + 1,23) &"<BR>"
        response.write tmpRec(CurrentPage,index + 1,24) &"<BR>" 
end  select
Session("YECE0901F") = tmpRec
%>
</html>

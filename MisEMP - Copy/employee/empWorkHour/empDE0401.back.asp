<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<%
SELF = "empDe0401"

ftype = request("ftype")
code = request("code")
index=request("index")
CurrentPage = request("CurrentPage")

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

response.write  "CODESTR13=" & CODESTR13 &"<BR>"

tmpRec = Session("empde0401")
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
	case "A"
		sql="select * from view_empfile  where empid='"& code &"'  "
		response.write sql &"<br>"
		'response.end
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,1) = Ucase(code)
  			tmpRec(CurrentPage,index + 1,7) = rst("groupid")
  			tmpRec(CurrentPage,index + 1,8) = rst("gstr")
  			tmpRec(CurrentPage,index + 1,10) = rst("empnam_cn")
  			tmpRec(CurrentPage,index + 1,11) = rst("empnam_vn")
  			tmpRec(CurrentPage,index + 1,12) = rst("whsno")  			
	  		empname=rst("empnam_cn") & rst("empnam_vn") &"-"&rst("nindat")
	  		tmpRec(CurrentPage,index + 1,15) = empname 
%>			<script language=vbs>
				Parent.Fore.<%=self%>.empid(<%=index%>).value="<%=Ucase(Code)%>"
				Parent.Fore.<%=self%>.empname(<%=index%>).value="<%=empname%>"
				Parent.Fore.<%=self%>.gstr(<%=index%>).value="<%=rst("gstr")%>"
				Parent.Fore.<%=self%>.groupid(<%=index%>).value="<%=rst("groupid")%>"
				Parent.Fore.<%=self%>.whsno(<%=index%>).value="<%=rst("whsno")%>"
				Parent.Fore.<%=self%>.wdat(<%=index%>).focus() 
				Parent.Fore.<%=self%>.wdat(<%=index%>).select() 
				'parent.best.cols="100%,0%"
			</script>
			
<% 		else%>
			<script language=vbs>
				alert "員工編號輸入錯誤!!"
				Parent.Fore.<%=self%>.empid(<%=index%>).value=""				
				Parent.Fore.<%=self%>.empid(<%=index%>).focus() 
				'parent.best.cols="100%,0%"
			</script>
<%		end if
		set rs=nothing
		'response.end 
	case "B"
		sql="select * from empwork where empid='"& CODESTR01 &"'  and workdat='"& replace(CODESTR02,"/","") &"'  " 
		response.write sql		
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3
  		if not rst.eof then
  			tmpRec(CurrentPage,index + 1,0) = "UPD"  			
  			tmpRec(CurrentPage,index + 1,13) = rst("timeup")
  			tmpRec(CurrentPage,index + 1,14) = rst("timedown")
%>			<script language=vbs>
				Parent.Fore.<%=self%>.b_timeup(<%=index%>).value="<%=rst("timeup")%>"
				Parent.Fore.<%=self%>.b_timedown(<%=index%>).value="<%=rst("timedown")%>"			
				'parent.best.cols="100%,0%"	
			</script>
<% 		end if
		set rs=nothing
	case "C"
		RESPONSE.WRITE "XXX"&"<BR>"
		tmpRec(CurrentPage,index + 1,0) = "UPD"
  		tmpRec(CurrentPage,index + 1,1) = CODESTR01
  		tmpRec(CurrentPage,index + 1,12) = CODESTR02
  		tmpRec(CurrentPage,index + 1,2) = CODESTR03
  		tmpRec(CurrentPage,index + 1,3) = CODESTR04
  		tmpRec(CurrentPage,index + 1,4) = CODESTR05
  		tmpRec(CurrentPage,index + 1,5) = CODESTR06
  		tmpRec(CurrentPage,index + 1,6) = CODESTR07
%>	 
<%
 
end  select
response.write "1="& tmpRec(CurrentPage,index + 1,1) &"<BR>"
response.write "2="& tmpRec(CurrentPage,index + 1,2) &"<BR>"
response.write "3="& tmpRec(CurrentPage,index + 1,3) &"<BR>"
response.write "4="& tmpRec(CurrentPage,index + 1,4) &"<BR>"
response.write "5="& tmpRec(CurrentPage,index + 1,5) &"<BR>"
response.write "6="& tmpRec(CurrentPage,index + 1,6) &"<BR>"
Session("empde0401") = tmpRec
%>
</html>

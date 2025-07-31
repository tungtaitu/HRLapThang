<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
SELF = "empBHGT" 

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
workdays = REQUEST("days")  
response.write  "workdays=" & workdays &"<BR>"

tmpRec = Session("empBHGTD") 
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
		sql="select * from empsalaryBasic where func='AA' and country='VN' and code='"& code &"'  "
		response.write sql
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,19) = code
  			tmpRec(CurrentPage,index + 1,20) = rst("bonus")
  			tmpRec(CurrentPage,index + 1,21) = cdbl(rst("bonus"))*0.05
  			tmpRec(CurrentPage,index + 1,22) = cdbl(rst("bonus"))*0.01  	  		
	  		tmpRec(CurrentPage,index + 1,29) = CODESTR06
	  		tmpRec(CurrentPage,index + 1,30) = CODESTR07
				tmpRec(CurrentPage,index + 1,32) = cdbl(rst("bonus"))*0.01  
				tmpRec(CurrentPage,index + 1,28) = cdbl(rst("bonus"))*0.05 + cdbl(rst("bonus"))*0.01+ cdbl(rst("bonus"))*0.01   
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>				
				Parent.Fore.<%=self%>.BHXH(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,21)%>
				Parent.Fore.<%=self%>.BHYT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,22)%>
				Parent.Fore.<%=self%>.BHTN(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,32)%>
				Parent.Fore.<%=self%>.BHTOT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,28)%>
				
			</script>
<% 		end if  
		set rs=nothing 		 
	case "CDATACHG1"		
		tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
  		tmpRec(CurrentPage,index + 1,19) = CODESTR02
  		tmpRec(CurrentPage,index + 1,20) = CODESTR01
  		tmpRec(CurrentPage,index + 1,21) = CODESTR03
  		tmpRec(CurrentPage,index + 1,22) = CODESTR04
  		tmpRec(CurrentPage,index + 1,23) = CODESTR05   		  		
  		tmpRec(CurrentPage,index + 1,29) = CODESTR06
	  	tmpRec(CurrentPage,index + 1,30) = CODESTR07
			tmpRec(CurrentPage,index + 1,32) = CODESTR08
			tmpRec(CurrentPage,index + 1,28) = cdbl(CODESTR03)+ cdbl(CODESTR04)+cdbl(CODESTR08)
%>		<script language=vbs>																				 						
			Parent.Fore.<%=self%>.BHXH(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,21)%>
			Parent.Fore.<%=self%>.BHYT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,22)%>
			Parent.Fore.<%=self%>.BHTN(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,32)%>
			Parent.Fore.<%=self%>.BHTOT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,28)%>
			Parent.Fore.<%=self%>.GTAMT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,23)%>
			'Parent.Fore.<%=self%>.BHYT(<%=index%>).FOCUS()
			Parent.Fore.<%=self%>.BHYT(<%=index%>).SELECT()
		</script>
<% 	
	case "CDATACHG"		
		tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
  		tmpRec(CurrentPage,index + 1,19) = CODESTR02
  		tmpRec(CurrentPage,index + 1,20) = CODESTR01
  		tmpRec(CurrentPage,index + 1,21) = CODESTR03
  		tmpRec(CurrentPage,index + 1,22) = CODESTR04
  		tmpRec(CurrentPage,index + 1,23) = CODESTR05   		  		
  		tmpRec(CurrentPage,index + 1,29) = CODESTR06
	  	tmpRec(CurrentPage,index + 1,30) = CODESTR07
			tmpRec(CurrentPage,index + 1,32) = CODESTR08
			tmpRec(CurrentPage,index + 1,28) = cdbl(CODESTR03)+ cdbl(CODESTR04)+ cdbl(CODESTR08)
%>		<script language=vbs>																				
			Parent.Fore.<%=self%>.BHXH(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,21)%>
			Parent.Fore.<%=self%>.BHYT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,22)%>
			Parent.Fore.<%=self%>.BHTN(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,32)%>
			Parent.Fore.<%=self%>.BHTOT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,28)%>
			Parent.Fore.<%=self%>.GTAMT(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,23)%>
		</script>
<%  		
end  select   		
Session("empBHGTD") = tmpRec
%>
</html>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
SELF = "empsalary01" 

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


tmpRec = Session("empsalary01") 
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
		sql="select * from empsalaryBasic where func='AA' and code='"& code &"'  "
		response.write sql &"<BR>"
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,19) = code
  			tmpRec(CurrentPage,index + 1,20) = rst("bonus")  			
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  			if tmpRec(CurrentPage,index + 1, 4) = "VN" then 
	  			TTMH = round( (TTM/26/8) , 0) 	  			 
	  		else
	  			TTMH = round(cdbl(rst("bonus"))/30/8 , 3 ) 
	  		end if 	
  			tmpRec(CurrentPage,index + 1,31)=TTMH 
  			
  			'+(加項)
	  		BB = tmpRec(CurrentPage,index + 1,20)
	  		CV = tmpRec(CurrentPage,index + 1,22)
	  		PHU = tmpRec(CurrentPage,index + 1,23)
	  		NN = tmpRec(CurrentPage,index + 1,24)
	  		KT = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TTKH = tmpRec(CurrentPage,index + 1,27)	  		
	  		QC = tmpRec(CurrentPage,index + 1,32)
	  		 
		  	F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) 
	  		
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>				
				Parent.Fore.<%=self%>.totamt(<%=index%>).value=<%=F1_MONEY%> 
			</script>
<% 		end if  
		set rs=nothing 		
	case "B"		
		sql="select * from empsalaryBasic where func='BB' and JOB='"& code &"'  "
		response.write sql &"<BR>" 
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,6) = code
  			tmpRec(CurrentPage,index + 1,21) = rst("CODE")
  			tmpRec(CurrentPage,index + 1,22) = rst("bonus")
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  			if TTM mod (26*8)<>0 then 
  				TTMH = fix(TTM/26/8)+1 
  			else
  				TTMH = fix(TTM/26/8) 
  			end if 
  			tmpRec(CurrentPage,index + 1,31)= TTMH     
  			
%>			<script language=vbs>												
				Parent.Fore.<%=self%>.CV(<%=index%>).value="<%=rst("bonus")%>"
				Parent.Fore.<%=self%>.CVCODE(<%=index%>).value="<%=rst("CODE")%>"
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value="<%=TTMH%>"
			</script>
<% 		end if  
		set rs=nothing 	
		sql2="select * from empsalaryBasic where func='CC' and JOB='"& code &"' "
		response.write sql2 &"<BR>" 
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL2, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,31) = rst("bonus")		 
  			
  			'+(加項)
	  		BB = tmpRec(CurrentPage,index + 1,20)
	  		CV = tmpRec(CurrentPage,index + 1,22)
	  		PHU = tmpRec(CurrentPage,index + 1,23)
	  		NN = tmpRec(CurrentPage,index + 1,24)
	  		KT = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TTKH = tmpRec(CurrentPage,index + 1,27)	  		
	  		QC_MONEY = rst("bonus") 		  
	  		
	  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC_MONEY)   

%>			<script language=vbs>												
				Parent.Fore.<%=self%>.QC(<%=index%>).value="<%=QC_MONEY%>"
				Parent.Fore.<%=self%>.totamt(<%=index%>).value="<%=F1_MONEY%>"
			</script>
<% 		end if   
		set rs=nothing   
		
	case "CDATACHG"		
		tmpRec(CurrentPage,index + 1,0) = "UPD"		   		
  		tmpRec(CurrentPage,index + 1,23) = CODESTR01
  		tmpRec(CurrentPage,index + 1,24) = CODESTR02
  		tmpRec(CurrentPage,index + 1,25) = CODESTR03
  		tmpRec(CurrentPage,index + 1,26) = CODESTR04
  		tmpRec(CurrentPage,index + 1,27) = CODESTR05
  		tmpRec(CurrentPage,index + 1,32) = CODESTR06
  		tmpRec(CurrentPage,index + 1,20) = CODESTR07
  		tmpRec(CurrentPage,index + 1,22) = CODESTR08
  		
  		TTM = cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  		if TTM mod (26*8)<>0 then 
  			TTMH = fix(TTM/26/8)+1 
  		else
  			TTMH = fix(TTM/26/8) 
  		end if 
  		tmpRec(CurrentPage,index + 1,31)=TTMH    
  		
  		'+(加項)
  		BB = tmpRec(CurrentPage,index + 1,20)
  		CV = tmpRec(CurrentPage,index + 1,22)
  		PHU = tmpRec(CurrentPage,index + 1,23)
  		NN = tmpRec(CurrentPage,index + 1,24)
  		KT = tmpRec(CurrentPage,index + 1,25)
  		MT = tmpRec(CurrentPage,index + 1,26)
  		TTKH = tmpRec(CurrentPage,index + 1,27)
  		QC = tmpRec(CurrentPage,index + 1,32)
  		 
  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC) 		
	 	response.write 	F1_MONEY &"<BR>"
  		tmpRec(CurrentPage,index + 1,33) = F1_MONEY 
%>		<script language=vbs>														
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			Parent.Fore.<%=self%>.totamt(<%=index%>).value=<%=(F1_MONEY)%> 
		'	alert <%=TTMH%>
		</script>
<%
end  select   		
Session("empsalary01") = tmpRec
%>
</html>

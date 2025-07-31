<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" --> 
<%
SELF = "YECE0101" 
SESSION.CODEPAGE=65001
ftype = request("ftype") 
code = request("code") 
lncode = request("lncode")
whsno = request("whsno")
index=request("index")  
CurrentPage = request("CurrentPage") 
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

CODESTR10wp = REQUEST("CODESTR10wp")


tmpRec = Session("empsalary01") 
response.write "index=" & index &"<BR>"
response.write "ftype=" & ftype &"<BR>"
response.write "whsno=" & whsno &"<BR>"
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
		if  tmpRec(CurrentPage,index + 1,4)="VN" then 
			sql="select * from empsalaryBasic where func='AA' and code='"& code &"' "&_
				"and country='"& tmpRec(CurrentPage,index + 1,4) &"' and bwhsno='"& whsno &"'"
		else
			sql="select * from empsalaryBasic where func='AA' and code='"& code &"' "&_
				"and country='"& tmpRec(CurrentPage,index + 1,4) &"'  "
		end if 	
		response.write sql &"<BR>"
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,39) = code
			tmpRec(CurrentPage,index + 1,47) = lncode
  			tmpRec(CurrentPage,index + 1,20) = rst("bonus")   
			response.write  "xxxx" & rst("bonus")    
				 
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  			if tmpRec(CurrentPage,index + 1, 4) = "VN" then 
	  			TTMH = round( (TTM/26/8) , 0) 	  			 
	  		else
	  			TTMH = round(cdbl(rst("bonus"))/30/8 , 3 ) 
	  		end if 	
  			tmpRec(CurrentPage,index + 1,31)=TTMH 
  			
  			'+(加項)
			if tmpRec(CurrentPage,index + 1,19)="" then wp=0 else  wp = tmpRec(CurrentPage,index + 1,19)
	  		if tmpRec(CurrentPage,index + 1,20)="" then bb= 0 else BB = tmpRec(CurrentPage,index + 1,20)
	  		if tmpRec(CurrentPage,index + 1,22)="" then cv= 0 else CV = tmpRec(CurrentPage,index + 1,22)
	  		if tmpRec(CurrentPage,index + 1,23)="" then phu = 0 else  PHU = tmpRec(CurrentPage,index + 1,23)
	  		if tmpRec(CurrentPage,index + 1,24)="" then nn=0 else  NN = tmpRec(CurrentPage,index + 1,24)
	  		if tmpRec(CurrentPage,index + 1,25)="" then kt = 0 else  KT = tmpRec(CurrentPage,index + 1,25)
	  		if tmpRec(CurrentPage,index + 1,26)="" then mt = 0 else  MT = tmpRec(CurrentPage,index + 1,26)
	  		if tmpRec(CurrentPage,index + 1,27)="" then ttkh = 0 else  TTKH = tmpRec(CurrentPage,index + 1,27)	  		
	  		if tmpRec(CurrentPage,index + 1,32)="" then qc = 0 else  QC = tmpRec(CurrentPage,index + 1,32)
			if tmpRec(CurrentPage,index + 1,44)="" then btien = 0 else  btien = tmpRec(CurrentPage,index + 1,44)
			if tmpRec(CurrentPage,index + 1,46)="" then btien3 = 0 else  btien3 = tmpRec(CurrentPage,index + 1,46)
			'if tmpRec(CurrentPage,index + 1,45)="" then tien2 = 0 else  tien2 = tmpRec(CurrentPage,index + 1,45)				
	  		'response.write (BB) & (CV) & (PHU) & (NN) & (KT) & (MT) & (TTKH) & (QC) & (wp)  
			'response.end
		  	F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+cdbl(wp)+cdbl(btien) 
	  		IF WHSNO="DN" THEN
				if cdbl(tmpRec(CurrentPage,index + 1,40))> 0 then 
					tmpRec(CurrentPage,index + 1,41)= cdbl(tmpRec(CurrentPage,index + 1,40)) - CDBL(BB) - CDBL(CV) 						
				end if 
				if cdbl(tmpRec(CurrentPage,index + 1,41)) >0 then 
					tmpRec(CurrentPage,index + 1,23)= tmpRec(CurrentPage,index + 1,41)
				end if 
			end if 	
%>			<script language=vbs>								
				Parent.Fore.<%=self%>.bb(<%=index%>).value=<%=rst("bonus")%>
				Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=TTMH%>				
				Parent.Fore.<%=self%>.totamt(<%=index%>).value=<%=F1_MONEY%> 
				Parent.Fore.<%=self%>.CBdiff(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,41)%> 
				Parent.Fore.<%=self%>.phu(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,23)%> 
			</script>
<% 		end if  
		set rs=nothing 		
	case "B"		
		if tmpRec(CurrentPage,index + 1,4)="VN" then 
			sql="select * from empsalaryBasic where func='BB' and JOB='"& code &"' "&_
				"and country='"& tmpRec(CurrentPage,index + 1,4) &"' and bwhsno='"& whsno &"' "
		else
			sql="select * from empsalaryBasic where func='BB' and JOB='"& code &"' "&_
				"and country='"& tmpRec(CurrentPage,index + 1,4) &"'  "		
		end if
		response.write sql &"<BR>" 
		'response.end  
		tmpRec(CurrentPage,index + 1,6) = code
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"  			
  			tmpRec(CurrentPage,index + 1,21) = rst("CODE")
  			tmpRec(CurrentPage,index + 1,22) = rst("bonus")
  			TTM = cdbl(rst("bonus"))+cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  			if TTM mod (26*8)<>0 then 
  				TTMH = fix(TTM/26/8)+1 
  			else
  				TTMH = fix(TTM/26/8) 
  			end if 
  			tmpRec(CurrentPage,index + 1,31)= TTMH      
  			c1 = tmpRec(CurrentPage,index + 1,4) 
  			
%>			<script language=vbs>												
				Parent.Fore.<%=self%>.CV(<%=index%>).value="<%=rst("bonus")%>"
				Parent.Fore.<%=self%>.CVCODE(<%=index%>).value="<%=rst("CODE")%>" 
			 	c1 = "<%=c1%>"
			  	if c1="VN" then 
					Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value="<%=TTMH%>"
				end if 	
			</script>
<% 		end if  
		set rs=nothing 	
		sql2="select * from empsalaryBasic where func='CC' and JOB='"& code &"' "&_
			 "and country='"& tmpRec(CurrentPage,index + 1,4) &"' and bwhsno='"& whsno &"' "
		response.write sql2 &"<BR>" 
		'response.end  
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL2, conn, 3,3      	
  		if not rst.eof then  
  			tmpRec(CurrentPage,index + 1,0) = "UPD"
  			tmpRec(CurrentPage,index + 1,32) = rst("bonus")		 
  			
  			'+(加項)
				wp = tmpRec(CurrentPage,index + 1,19)
	  		BB = tmpRec(CurrentPage,index + 1,20)
	  		CV = tmpRec(CurrentPage,index + 1,22)
	  		PHU = tmpRec(CurrentPage,index + 1,23)
	  		NN = tmpRec(CurrentPage,index + 1,24)
	  		KT = tmpRec(CurrentPage,index + 1,25)
	  		MT = tmpRec(CurrentPage,index + 1,26)
	  		TTKH = tmpRec(CurrentPage,index + 1,27)	  		
				btien = tmpRec(CurrentPage,index + 1,44)	  		
				'tien2 = tmpRec(CurrentPage,index + 1,45)	  	
	  		QC_MONEY = rst("bonus") 		  
	  		
	  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC_MONEY)+cdbl(wp)+cdbl(btien)   
				IF WHSNO="DN" THEN
					if cdbl(tmpRec(CurrentPage,index + 1,40))> 0 then 
						tmpRec(CurrentPage,index + 1,41)= cdbl(tmpRec(CurrentPage,index + 1,40)) - CDBL(BB) - CDBL(CV) 
					end if 
					if cdbl(tmpRec(CurrentPage,index + 1,41)) >0 then 
						tmpRec(CurrentPage,index + 1,23)= tmpRec(CurrentPage,index + 1,41)
					end if 
				end if 	
%>			<script language=vbs>												
				Parent.Fore.<%=self%>.QC(<%=index%>).value="<%=QC_MONEY%>"
				Parent.Fore.<%=self%>.totamt(<%=index%>).value="<%=(F1_MONEY)%>"
				Parent.Fore.<%=self%>.CBdiff(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,41)%> 
				Parent.Fore.<%=self%>.phu(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,23)%> 
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
  		tmpRec(CurrentPage,index + 1,34) = CODESTR09
		tmpRec(CurrentPage,index + 1,19) = CODESTR10
		tmpRec(CurrentPage,index + 1,44) = cdbl(CODESTR11) 
		tmpRec(CurrentPage,index + 1,45) = CODESTR12
		tmpRec(CurrentPage,index + 1,46) = CODESTR10wp
  		
  		TTM = cdbl(tmpRec(CurrentPage, index+1, 20))+cdbl(tmpRec(CurrentPage, index+1, 22))+cdbl(tmpRec(CurrentPage, index+1, 23)) 
  		if TTM mod (26*8)<>0 then 
  			TTMH = fix(TTM/26/8)+1 
  		else
  			TTMH = fix(TTM/26/8) 
  		end if 
  		tmpRec(CurrentPage,index + 1,31)=TTMH    
  		
  		'+(加項)
		wp = tmpRec(CurrentPage,index + 1,19)
  		BB = tmpRec(CurrentPage,index + 1,20)
  		CV = tmpRec(CurrentPage,index + 1,22)
  		PHU = tmpRec(CurrentPage,index + 1,23)
  		NN = tmpRec(CurrentPage,index + 1,24)
  		KT = tmpRec(CurrentPage,index + 1,25)
  		MT = tmpRec(CurrentPage,index + 1,26)
  		TTKH = tmpRec(CurrentPage,index + 1,27)
  		QC = tmpRec(CurrentPage,index + 1,32)
		btien = tmpRec(CurrentPage,index + 1,44)
		wpbtien = tmpRec(CurrentPage,index + 1,46)
		'tien2 = tmpRec(CurrentPage,index + 1,45)
  		 
  		F1_MONEY = CDBL(BB)+CDBL(CV)+CDBL(PHU)+CDBL(NN)+CDBL(KT)+CDBL(MT)+CDBL(TTKH)+CDBL(QC)+cdbl(wp)+cdbl(btien)+cdbl(wpbtien)
			IF WHSNO="DN" THEN
				if cdbl(tmpRec(CurrentPage,index + 1,40))> 0 then 
					tmpRec(CurrentPage,index + 1,41)= cdbl(tmpRec(CurrentPage,index + 1,40)) - CDBL(BB) - CDBL(CV) 
				end if 
				if cdbl(tmpRec(CurrentPage,index + 1,41))>0 then 
					tmpRec(CurrentPage,index + 1,23) = tmpRec(CurrentPage,index + 1,41)
					phu = tmpRec(CurrentPage,index + 1,41)
				end if 
			end if 				 
				
	 	response.write 	F1_MONEY &"<BR>"
  		tmpRec(CurrentPage,index + 1,33) = F1_MONEY 
%>		<script language=vbs>														
			Parent.Fore.<%=self%>.phu(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,23)%> 
			Parent.Fore.<%=self%>.HHMOENY(<%=index%>).value=<%=(TTMH)%>
			Parent.Fore.<%=self%>.totamt(<%=index%>).value="<%=formatnumber(F1_MONEY,0)%>"
			Parent.Fore.<%=self%>.CBdiff(<%=index%>).value=<%=tmpRec(CurrentPage,index + 1,41)%>			
		'	alert <%=TTMH%>
		</script>
<%
end  select   		 
response.write "wp=" & tmpRec(CurrentPage,index + 1,19) &"<BR>" 
response.write "job=" & tmpRec(CurrentPage,index + 1,6) &"<BR>"
response.write "bb=" & tmpRec(CurrentPage,index + 1,20) &"<BR>"
response.write "cv=" & tmpRec(CurrentPage,index + 1,22) &"<BR>" 
response.write "phu=" & tmpRec(CurrentPage,index + 1,23) &"<BR>"
response.write "nn=" & tmpRec(CurrentPage,index + 1,24) &"<BR>"
response.write "kt=" & tmpRec(CurrentPage,index + 1,25) &"<BR>"
response.write "mt=" & tmpRec(CurrentPage,index + 1,26) &"<BR>"
response.write "ttkh=" & tmpRec(CurrentPage,index + 1,27) &"<BR>"
response.write "qc=" & tmpRec(CurrentPage,index + 1,32) &"<BR>"
response.write "tot=" & tmpRec(CurrentPage,index + 1,33) &"<BR>"
response.write "memo=" & tmpRec(CurrentPage,index + 1,34) &"<BR>" 
response.write "btien=" & tmpRec(CurrentPage,index + 1,44) &"<BR>" 
'response.write "tien2=" & tmpRec(CurrentPage,index + 1,45) &"<BR>" 
Session("empsalary01") = tmpRec
%>
</html>


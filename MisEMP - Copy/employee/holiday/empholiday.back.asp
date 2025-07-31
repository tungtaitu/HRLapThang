<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" --> 
<%
SELF = "empholiday" 

ftype = request("ftype") 
code = request("code")  
code1=request("code1")
Set conn = GetSQLServerConnection()	  
yymm= year(date())&right("00"&month(date()),2)    '本月 
nowmonth = year(date())&right("00"&month(date()),2)    '本月 
calcdt = left(YYMM,4)&"/"& right(yymm,2)&"/01"    '本月第1天
 '一個月有幾天
cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01")
days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
'本月最後一天
ENDdat = CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/"&DAYS)
ENDdat=year(ENDdat)&"/"&right("00"&month(Enddat),2)&"/"&right("00"&day(Enddat),2) 

'下個月
if right(yymm,2)="12" then
	ccdt = cstr(left(YYMM,4)+1)&"/01/01"
else
	ccdt = left(YYMM,4)&"/"& right("00" & right(yymm,2)+1,2)  &"/01"
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>
<%
select case ftype 
	case "chkempid"		
		sql="select * from view_empfile  where empid='"& code &"' "
		response.write sql&"<br>"
		'response.end   
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3      		
  		if not rst.eof then    		
			if trim(rst("outdate"))="" then 
				tx_enddat = ccdt	
			else			
				if right(rst("outdate"),5)>= "03/31"  and  year(rst("indat")) < year(date()) then
					'tx_enddat = cstr(left(yymm,4))&"/04/01"
					tx_enddat = rst("outdate") 
				else
					tx_enddat = rst("outdate") 
				end if			
			end if 	
				
			if  right(yymm,2)<="03" and yymm<=nowmonth     then 
				response.write "1"&"<br>"
				if cdate( rst("calcTxdat") ) <= cdate(cstr(left(yymm,4)-1)&"/03/31") then 
					cc_indat = cstr(left(yymm,4)-1)&"/04/01" 
				else
					cc_indat = year(rst("calcTxdat"))&"/"&right("00"&month(rst("calcTxdat")),2)&"/"&right("00"&day(rst("calcTxdat")),2)
				end if	
			else				
				response.write "2"	&"<br>"
				if year(rst("indat")) >=  year(date()) then 
					response.write "21"	&"<br>"
					cc_indat = year(rst("calcTxdat"))&"/"&right("00"&month(rst("calcTxdat")),2)&"/"&right("00"&day(rst("calcTxdat")),2)
				else
					cc_indat = cstr(left(yymm,4))&"/04/01" 		
				end if 	
			end if   
			response.write "cc_indat=" & cc_indat &"<BR>"
			if rst("outdate")<= year(rst("calcTxdat"))&"/"&right("00"&month(rst("calcTxdat")),2)&"/"&right("00"&day(rst("calcTxdat")),2) then 
				cc_indat = rst("nindat") 
			end if 
			year_tx = datediff("m", cdate(cc_indat) , cdate(tx_enddat) ) *8 
			
			
			if rst("outdate") = "" then 
				sqlx="select  empid , sum(hhour) hhour  from empholiday where jiatype='E' and empid='"& code &"'  "&_
					 "and convert(char(10),dateup,111) between '"& cc_indat &"' and  convert(char(10), getdate(),111)   "&_
					"group by empid "
			else
				sqlx="select  empid , sum(hhour) hhour  from empholiday where jiatype='E' and empid='"& code &"'  "&_
					 "and convert(char(10),dateup,111) between '"& cc_indat &"' and '"& tx_enddat &"'   "&_			
					 "group by empid "
			end if 
			set rsx=conn.execute(Sqlx) 
			if rsx.eof then 
				hhour = 0 
			else
				hhour = rsx("hhour") 
			end if 
			set rsx=nothing 
			response.write sqlx &"<br>" 
			
			datTD = cc_indat & "~"&tx_enddat 	 
			yy = cdbl(year_tx)-cdbl(hhour) 
%>			<script language=vbs>				
					Parent.Fore.<%=self%>.empid.value = "<%=rst("empid")%>"
			    Parent.Fore.<%=self%>.nvx.value = "<%=rst("nindat")%>"
					otyn="<%=rst("outdate")%>" 
					if otyn<>"" then 
						Parent.Fore.<%=self%>.ntv.value = " NTV:"&"<%=rst("outdate")%>"
					else	
						Parent.Fore.<%=self%>.ntv.value = "<%=rst("outdate")%>"
					end if 	
			    Parent.Fore.<%=self%>.EMPNAMEVN.value = "<%=rst("empnam_CN")%>"&"<%=rst("empnam_VN")%>" 
					Parent.Fore.<%=self%>.country.value="<%=rst("country")%>"
					cn = "<%=rst("country")%>" 
					if cn="VN" then 
						Parent.Fore.<%=self%>.radio1(0).checked=true  
						Parent.Fore.<%=self%>.place.value=""	
					end if 
			    Parent.Fore.<%=self%>.ynjh.value = "<%=yy%>" 
					nyy="<%=yy%>"
					if cdbl(nyy) < 0 then 
						Parent.Fore.<%=self%>.ynjh.style.color="red"
					else	
						Parent.Fore.<%=self%>.ynjh.style.color="black"
					end if 
					Parent.Fore.<%=self%>.ynj.value = "<%=datTD%>" 
					Parent.Fore.<%=self%>.HOLIDAY_TYPE.focus()
			</script>
<%  					
  		else   			
%>			<script language=vbs>				
				alert "員工編號輸入錯誤!!"
				Parent.Fore.<%=self%>.empid.focus()
				Parent.Fore.<%=self%>.empid.value = ""			    
			  Parent.Fore.<%=self%>.EMPNAMEVN.value = ""
			</script>
<%			response.end   			 
 		end if  
		set rs=nothing  
		
	case "dayschg"
		sql="select  isnull(count(*),0)   as ccnt  from   ydbmcale  "&_
			"where status in ( 'H2', 'H3' )  "&_
			"and convert(char(10), dat,111) between '"& code &"' and '"& code1 &"' " 
		response.write sql 	
		Set rst= Server.CreateObject("ADODB.Recordset")
  		rst.Open SQL, conn, 3,3  
		if not rst.eof then    			 
%>			<script language=vbs>				
				Parent.Fore.<%=self%>.HDcnt.value = "<%=rst("ccnt")%>"			    
			</script>
<%  					
  		else   			
%>			<script language=vbs>				
				Parent.Fore.<%=self%>.HDcnt.value="0"				
			</script>
<%			response.end   			 
 		end if  
		set rs=nothing    		
		
		 	
end  select   		
%>

</html>

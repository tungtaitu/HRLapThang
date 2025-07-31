<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<%
'on error resume next   
session.codepage="65001"
SELF = "yeee03"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
yymm = REQUEST("yymm")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
empid1 = trim(request("empid1"))  
sortby = trim(request("showby"))  

gTotalPage = 1
PageRec = 0    'number of records per page
TableRec = 60    'number of fields per record  

tjnum = 12 
sqlx="select datediff(m,td1, td2)+1 as tjnum from empNJYM_set where njym='"&yymm&"' "
set rsx=conn.execute(Sqlx)
if not rsx.eof then 
	tjnum = rsx("tjnum")
end if 
set rsx=nothing 
'response.write   tjnum 
'response.end 

sql="exec SP_empNJTJ_N '"& yymm &"', '"& country &"','"& whsno &"','"& groupid &"','"& empid1 &"','"& sortby &"' "  
'response.write sql 
'RESPONSE.END  
'if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 3 
	IF NOT RS.EOF THEN 
		while not rs.eof  
			PageRec= PageRec + 1 
			rs.movenext
		wend 
		rs.PageSize = PageRec 
		RecordInDB = PageRec 
		TotalPage = 1
		gTotalPage = TotalPage
		rs.movefirst
	END IF 	 
	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
	 for j = 1 to PageRec
		if not rs.EOF then 			
				td1= rs("td1")
				td2= rs("td2")
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("empid"))
				tmpRec(i, j, 2) = trim(rs("empnam_cn"))
				tmpRec(i, j, 3) = trim(rs("empnam_vn"))
				tmpRec(i, j, 4) = rs("country")
				tmpRec(i, j, 5) = rs("nindat")
				tmpRec(i, j, 6) = rs("outdate")
				tmpRec(i, j, 7) = rs("job")				
				tmpRec(i, j, 8) = rs("whsno")	 				 
				tmpRec(i, j, 9)	=RS("groupid") 
				tmpRec(i, j, 10)=RS("gstr") 						
				tmpRec(i, j, 11)=RS("jstr") 	
				 

				tmpRec(i, j, 12)=RS("jiaE_hr")
				tmpRec(i, j, 13)=RS("tx_hr")
				tmpRec(i, j, 14)=RS("txdays")
				tmpRec(i, j, 15)=RS("nowtx")  '剩下特休時數
				tmpRec(i, j, 16)=RS("TD1")  '統計日期(起)
				tmpRec(i, j, 17)=RS("TD2")	'統計日期(迄)
				tmpRec(i, j, 18)=RS("txdat_s") 	'員工特休統計日期(起)
				tmpRec(i, j, 19)=RS("txdat_e")	'員工特休統計日期(迄)				
				tmpRec(i, j, 20)=RS("hh_money")	'平均時薪
				tmpRec(i, j, 21)=RS("min_dat")	'產假(起)
				tmpRec(i, j, 22)=RS("max_dat")	'產假(迄)
				tmpRec(i, j, 23)=RS("jiaF")	'修產假共?月				
				tmpRec(i, j, 24)=rs("TXmemo")   'memo
				tmpRec(i, j, 25)=rs("tz_txd")   '調整天數
				tmpRec(i, j, 26)=rs("nj_amt")   '未修完年假代金 
				
				for kk = 1 to tjnum 
					tmpRec(i, j, 26+kk) = rs(4+kk)  '' 每月特休時數
					tmpRec(i, j, 26+tjnum+kk) = rs(4+kk).name    '欄位名稱 T103H ( 第2碼表示年  ,ex : 2011  就是1 , 2012 = 2 , 3.4碼為月 ) 
					'response.write 26+kk &" "&tmpRec(i, j, 26+kk) &" --- "& tmpRec(i, j, 26+tjnum+kk) &"<BR>"
				next 				
				
				rs.MoveNext  				
				'response.write  tmpRec(i, j, 1) &"<BR>" 
		else 
			exit for 
		end if 
	 next 
	
	 if rs.EOF then 
		rs.Close 
		Set rs = nothing
		exit for 
	 end if 
	next 
	
' else    
	' TotalPage = cint(request("TotalPage"))	
	' CurrentPage = cint(request("CurrentPage"))
	' RecordInDB  = REQUEST("RecordInDB")
	 
	' Select case request("send") 
	     ' Case "FIRST"
		      ' CurrentPage = 1			
	     ' Case "BACK"
		      ' if cint(CurrentPage) <> 1 then 
			     ' CurrentPage = CurrentPage - 1				
		      ' end if
	     ' Case "NEXT"
		      ' if cint(CurrentPage) <= cint(TotalPage) then 
			     ' CurrentPage = CurrentPage + 1 
		      ' end if			
	     ' Case "END"
		      ' CurrentPage = TotalPage 			
	     ' Case Else 
		      ' CurrentPage = 1	
	' end Select 
' end if   

'response.end 
FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)  	
END FUNCTION 

nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12" 
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	 

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh"> 
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  >
<%
  filenamestr = "emp_PN_"&minute(now)&second(now)&yymm&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
  
%> 
<table width=600 class="txt8"  cellspacing="1" cellpadding="1" border=1  >
	<tr BGCOLOR="LightGrey" height=22>		
		<TD align=center >STT</TD> 		
		<TD align=center >工號<br>so the</TD> 		
 		<TD align=center >姓名(中)<br>ho ten</TD>
		<TD align=center >姓名(越)<br>ho ten</TD>
		<TD align=center >到職日<br>NVX</TD>
		<TD align=center >離職日<br>NTC</TD>
		<TD align=center >職務<br>chu vu</TD>		
 		<TD align=center >單位ID</TD>
		<TD align=center >單位<br>Don vi</TD>
		<%for k1=1 to tjnum %>	
			<TD align=center width=30 nowrap ><%=left(tmprec(1,1,26+tjnum+k1),4)%><br>H</TD>
		<%next%>
		<TD align=center >產假(起)</TD>
		<TD align=center >產假(迄)</TD>
		<TD align=center >產假(月)</TD>
		<TD align=center >調整<br>(天)</TD>
		<TD align=center >年假<br>(天)</TD>
		<TD align=center >年假<br>(時數)</TD>
		<TD align=center >剩餘年假<br>(時數)</TD>		
		<%if session("rights")<="2" then %>
			<TD align=center width=50 nowrap >時薪<br>(平均)</TD> 
			<TD align=center width=70 nowrap >年假薪資</TD> 
		<%end if%>		
	</tr>
	 
	<%for x = 1 to PageRec 
	%>
	<TR   > 	
		<TD align=center ><%=x%></td>
		<TD ><%=trim(tmpRec(CurrentPage,x,1))%></td>
		<TD align=left ><%=tmprec(1,x,2)%></td><td><%=left(tmprec(1,x,3),22)%></td>
		<TD align=center ><%=tmprec(1,x,5)%></td>
		<td><font color="red"><%=tmprec(1,x,6)%></font>
		</td>
		<TD align=left ><%=left(tmprec(1,x,11),8)%></td>
		<TD align=left ><%=tmprec(1,x,9)%></td><td><%=tmprec(1,x,10)%></td>
		<%for y = 1 to tjnum %>
				<TD align=center ><%if tmprec(1,x,26+y)<>"0" then %><%=tmprec(1,x,26+y)%><%end if%></td>
		<%next%>
		
		<TD align=center ><%=tmprec(1,x,21)%></td>
		<TD align=center ><%=tmprec(1,x,22)%></td>		
		<TD align=center ><%if tmprec(1,x,23)>"0" then%><%=round(tmprec(1,x,23),1)%><%end if%></td>
		<%
			if cdbl(tmprec(1,x,25))=0 then 
				tz_txd="" 
				atz_txd = tmprec(1,x,14) 
				atz_txh = tmprec(1,x,13) 
				nowtxH = tmprec(1,x,15)
				if cdbl(tmprec(1,x,15))>0 then
					nj_amt = formatnumber(cdbl(tmprec(1,x,15))*3*cdbl(tmprec(1,x,20)),0)
				else
					nj_amt = 0 
				end if 	
			else
				tz_txd=tmprec(1,x,25)
				atz_txd = cdbl(tmprec(1,x,14)) + cdbl(tmprec(1,x,25))
				atz_txh = cdbl(atz_txd)*8.0
				nowtxH = cdbl(atz_txh) - cdbl(tmprec(1,x,12))
				nj_amt =  formatnumber(tmprec(1,x,26),0)
			end if 
		%>		
		<TD align=center ><%=tz_txd%></td>		
		<TD align=center ><%=atz_txd%></td>
		<TD align=center ><%=atz_txh%></td>		
		<TD align=center ><%=nowtxH%></td>				
		<%if session("rights")<="2" then %>
			<TD align=center><%=formatnumber(tmprec(1,x,20),0)%></td>		
			<TD align=right><% if cdbl(nowtxH)>0 then%><%=nj_amt%><%else%>0<%end if%></td>		
		<%end if%>		
	</TR>
	<%next%> 	  
</table>	
<%response.end %>  
  

</body>
</html>
 



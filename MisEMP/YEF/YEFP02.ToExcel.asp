<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->

<%

self="YEFP02"  


nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="1" then  
	calcmonth = year(date())-1&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="1" then  
		calcmonth = year(date())-1&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if  
yymm = trim(request("YYMM"))

code01=request("whsno")
code02="'"&replace(replace(request("groupid")," ","'"),",","',")&"'"
code03="'"&replace(replace(request("country")," ","'"),",","',")&"'"
'response.write request("country") &"<BR>" 
'response.write code03 &"<BR>" 

code04=request("JOB")
code05=trim(request("empid1"))
code06=trim(request("empid2"))
code07=trim(request("indat1"))
code08=trim(request("indat2"))
code09=request("empTJ")
code10=trim(request("bhdat1"))
code11=trim(request("bhdat2"))
code12=request("outemp")  
code13=request("zuno")   

sortby = request("orderby") 

if sortby="" then sortby="case when country='TW' then '0' else country end , country, empid" 

Set conn = GetSQLServerConnection()	  

sql="select a.*, isnull(b.whsno_acc,'') whsno_acc from "&_
	"( select * from view_empfile where empid<>'PELIN' and  isnull(empid,'')<>'' and whsno like '"& code01 &"%'   "&_
	"and zuno like '"& code13 &"%' and job like '"& job &"%' "&_
	"and convert(char(6),indat, 112)<='"& yymm &"' and empid like '%"& code05 &"%' "
if trim(request("country"))<>"" then 
	sql=sql & " and  country in ( " & code03 &" )" 
else
	sql=sql & " and  country  like '%' " 	
end if 	 	
if trim(request("groupid"))<>"" then 
	sql=sql & " and  groupid in ( " & code02 &" )"  
else	
	sql=sql & " and  groupid  like '%' " 	
end if 

if code07<>"" and code08<>"" then 
	sql=sql & " and convert(char(10), indat, 111) between '"& code07 &"' and '"& code08 &"'  "
end if 	
if code10<>"" and code11<>"" then 
	sql=sql & " and convert(char(10), bhdat, 111) between '"& code10 &"' and '"& code11 &"' "
end if 	
if code12="Y"  then 
	sql=sql & "and (isnull(outdat,'')='' or  convert(char(8),outdat,112)> '"&yymm&"01' )" 
elseif code12="N" then 
	sql=sql & "and isnull(outdat,'')<>'' and convert(varchar(6), outdat,112)<='"& yymm &"' "
end  if 	
sql=sql&" )a " 
sql=sql&"left join ( select empid as eid2 , whsno_acc  from empfile_acc ) b on b.eid2=a.empid  " 


sql=sql &" order by "& sortby  
'response.write sql  
'[Rpt_Empfile]
'@yymm as varchar(6),  --1
'@c1 as varchar(50),   --2
'@w1 as varchar(10), --3
'@G1 as varchar(60), --4
'@z1 as varchar(5), --5 
'@empid as varchar(10), --6
'@D1 as varchar(10), --7 
'@D2 as varchar(10), --8 
'@hD1 as varchar(10),
'@hD2 as varchar(10),
'@empTJ as varchar(1),   
'@outemp as varchar(1), 
'@sortby   as varchar(100),
'@otD1 as varchar(10),
'@otD2 as varchar(10)  --15 

'source  = "exec [Rpt_Empfile] '"&&"'  "

'set rs=conn.execute(Sql)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1,1
'response.end  
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
</head> 
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"   >
<%
  filenamestr = "empfile"&yymm&"_"&minute(now)&second(now)&".xls"
  Response.AddHeader "content-disposition","attachment; filename=" & filenamestr
  Response.Charset ="BIG5"
  Response.ContentType = "Content-Language;content=zh-tw" 
  Response.ContentType = "application/vnd.ms-excel"
%>
<span style="font-size:12pt"><font size=+1><b>CÔNG TY TNHN HOÀ ĐƯỜNG</b></font></span><br>
<span style="font-size:12pt"><font size=+1>員工資料明細表<%if code01<>"" then%>(<%=code01%>)<%end if %></font></span>
<TABLE CLASS="txt12" BORDER=1 cellspacing="1" cellpadding="1" style="font-size:9pt">	 
	<TR HEIGHT=25 BGCOLOR="#e4e4e4">
		<td align=center>STT</td>
		<td align=center>國籍</td>
		<td align=center>廠別</td>
		<td align=center>立帳</td>		
		<td align=center>工號</td>				
		<td align=center>姓名(中)</td>
		<td align=center>姓名(英越)</td>
		<td align=center>性別(M/F)</td>		
		<td align=center>到職日</td>		
		<td align=center>離職日</td>
		<td align=center>年資(月)</td>		
		<td align=center>班別</td>
		<td align=center>部門</td>
		<td align=center>單位</td>
		<td align=center>職碼</td>
		<td align=center>職稱</td>		
		<td align=center>學歷</td>		
		<td align=center>保險日期<br>(工作證號)</td>		
		<td align=center>出生日期</td>
		<td align=center>銀行帳號</td>
		<td align=center>身分証號</td>		
		<td align=center>發證地址<br>Nơi cấp</td>		
		<td align=center>護照號碼</td>
		<td align=center>護照簽發日</td>
		<td align=center>護照有效期</td>
		<td align=center>保險號碼<br>so bao hiem</td>
		<td align=center>address</td>		
		<td align=center>備註</td>		
	</tr>
	<%x = 0  
	  nianzi = 0	
	  while not rs.eof   	
			'response.write rs("empid") &"-"&rs("PDUEDATE") &"<BR>"
			x= x +1   
			nianzi = datediff("m", rs("nindat"),date()) 
	%> 	
			<TR HEIGHT=22 BGCOLOR="#ffffff">				
				<td align=left><%=x%></td>
				<td align=left><%=rs("country")%></td>
				<td align=left><%=rs("whsno")%></td>
				<td align=left><%=rs("whsno_acc")%>&nbsp;</td>				
				<td align=left><%=rs("empid")%></td>				
				<td align=left nowrap><%=rs("empnam_cn")%></td>
				<td align=left nowrap><%=Ucase(rs("empnam_vn"))%></td>
				<td align=left><%=rs("sex")%>&nbsp;</td>
				<td style="mso-number-format:\@" align=left><%=rs("nindat")%></td>
				<td style="mso-number-format:\@" align=left><%=rs("outdate")%></td>
				<td align=center><%=nianzi%></td>
				<td  align=left><%=rs("shift")%></td>
				<td  align=left><%=rs("gstr")%></td>
				<td  align=left><%=rs("zstr")%></td>
				<td  align=left><%=rs("job")%></td>
				<td  align=left><%=rs("jstr")%></td>
				<td  style="mso-number-format:\@" align=left><%=rs("school")%></td>				
				<td  align=left>&nbsp;<%=rs("bhdat")%></td>				
				<td  style="mso-number-format:\@" align=left><%=rs("bdy_ymd")%></td>
				<td  align=left>&nbsp;<%=rs("bankid_str")%></td>
				<td  align=left>&nbsp;<%=rs("personid")%></td>				
				<td style="mso-number-format:\@" align=left><%=rs("noiCap")%></td>
				<td style="mso-number-format:\@" align=left><%=rs("passportNo")%></td>
				<td style="mso-number-format:\@" align=left><%=rs("pissuedate")%></td>
				<td style="mso-number-format:\@" ><%=rs("passport_enddat")%></td>
				<td style="mso-number-format:\@" align=left><%=rs("sobh")%></td>				
				<td style="mso-number-format:\@" align=left><%=rs("homeaddr")%></td>				
				<td  align=left>&nbsp;<%=rs("memo")%></td>
			</tr> 
	<%		 
			rs.movenext
	  wend 	
		rs.close : set rs=nothing
		conn.close : set conn=nothing
	%>   	
</table> 
<%response.end%>

</body>
</html> 
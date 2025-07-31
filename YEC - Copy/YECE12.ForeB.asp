<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!-- #include file="../Include/SIDEINFO.inc" -->
<%
Set conn = GetSQLServerConnection()	  
self="yece12"  


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

yymm=request("yymm")
g1=request("g1")
qcountry=trim(request("qcountry"))
whsno=request("whsno")
eid=request("eid") 

nd1=left(yymm,4)&"/"&right(yymm,2)&"/01"
nd2=year(dateadd("m",1,nd1 ))&"/"&right("00"&month(dateadd("m",1,nd1 )),2)&"/01"

sqlx="select * from VYFYEXRT where  yyyymm='"& yymm &"' and code='USD'  "
set rdsx= conn.execute(sqlx)
if rdsx.eof then 
	response.write "本月匯率尚未建檔!!" 
	'response.end 
	rate = 1 
else	
	rate = rdsx("exrt")		
end if 	
rdsx.close : set rdsx=nothing 

If Request.ServerVariables("REQUEST_METHOD") = "POST" and request("btn")="GO" then 
	arr_empid=split(request("empid")&",",",")
	arr_money1=split(request("money1")&",",",")
	arr_country=split(request("country")&",",",")
	arr_money1usd=split(request("money1_usd")&",",",")
	'response.write arr_empid &"<BR>"
	'response.write arr_money1 &"<BR>"
	for k = 1 to ubound(arr_empid)
		
		f_empid=trim(arr_empid(k-1))
		f_money1=trim(replace( trim(arr_money1(k-1)) , ",",""))
		f_country=trim(replace( trim(arr_country(k-1)) , ",",""))
		f_money1usd=trim(replace( trim(arr_money1usd(k-1)) , ",",""))
		if f_money1<>"" then 
			'if f_country="VN" then money1=f_money1 else money1 = f_money1usd
			sql="if not exists ( select * from empmoney where yymm2='"&yymm&"' and empid2='"&f_empid&"'  ) "
			sql=sql&" insert into empmoney ( yymm2,empid2, money1 , mdtm,muser) values ( '"&yymm&"','"&f_empid&"','"&f_money1 &"' ,getdate(), '"&session("netuser")&"') "
			sql=sql&" else update empmoney set money1='"&f_money1&"' , mdtm=getdate(), muser='"&session("netuser")&"'  where yymm2='"&yymm&"' and empid2='"&f_empid&"' "
			conn.execute(sql)
			'response.write sql&"<BR>"
		end if 
	next 
end if 

sql="select a.* , isnull(b.money1,0) money1 from  (select * from view_empfile where nindat<'"& nd2 &"' and (isnull(outdat,'')='' or outdat>='"&nd1&"' ) "
sql=sql&" and  case when '"&qcountry&"'='' then '' else country end = '"&qcountry&"' "
sql=sql&" and  case when '"&whsno&"'='' then '' else whsno end = '"&whsno&"' "
sql=sql&" and  case when '"&g1&"'='' then '' else groupid end = '"&g1&"' "
sql=sql&" and  case when '"&eid&"'='' then '' else empid end = '"&eid&"' ) a "
sql=sql&" left join ( select * from empmoney where yymm2='"&yymm&"' ) b on b.empid2=a.empid "
sql=sql&" order by a.groupid, a.empid "
'response.write sql 
Set rs = Server.CreateObject("ADODB.Recordset")   
rs.open sql, conn, 3,1 
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<script language="JavaScript" src="enter2tab.js"></script>
</head> 
<body  topmargin="40" leftmargin="5"  marginwidth="0" marginheight="0"    >
<form name="foem1" id="from1" method="post"  action="yece12.foreb.asp">

<table width=100%  ><tr><td >
	<table  border=0 cellspacing="1" cellpadding="2"  class="txt8"> 
		<tr height=30 >
			<TD nowrap align=right  >年月<br>YYYYMM</TD>
			<TD ><INPUT NAME=yymm  CLASS=INPUTBOX VALUE="<%=yymm%>" SIZE=10 onchange="submit()"></TD>			
		 	<TD nowrap align=right height=30 >國籍<br>Quoc tich</TD>
			<TD >
				<select name=qcountry  class="txt8"  onchange="submit()" >					
					<%
					if request("CT")="VN" then 
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN'  ORDER BY SYS_type desc  "
					elseif request("CT")="CN" then 	
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='CN'  ORDER BY SYS_type desc  "
					elseif request("CT")="TA" then 	
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='TA'  ORDER BY SYS_type desc  "	
					elseif request("CT")="TM" then 	
						SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='X' ORDER BY SYS_type desc  "		
					else 	
						if session("rights")<="0" then 
							SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "		 
						else
							SQL="SELECT * FROM BASICCODE WHERE FUNC='country' and sys_type='VN' ORDER BY SYS_type desc  "		 
						end if 	
					end if 	
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if qcountry=RST("SYS_TYPE") then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>
					<%if request("CT")="TM"  then %>
					<option value="TM">TW+MA</option>
					<%end if%>
					<%if   session("rights")<="0" then %>
					<option value="HW">(hải ngoại)ALL海外</option>
					<%end if %>
				</SELECT>
				<%rst.close: SET RST=NOTHING %>
			</TD>			
			<TD nowrap align=right height=30 >廠別<br>Xuong</TD>
			<TD > 
				<select name=whsno  class=font9 onchange="submit()" >
					<option value="">--ALL---</option>
					<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
					SET RST = CONN.EXECUTE(SQL)
					WHILE NOT RST.EOF  
					%>
					<option value="<%=RST("SYS_TYPE")%>" <%if whsno=RST("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
					<%
					RST.MOVENEXT
					WEND 
					%>					
				</SELECT>
				<%rst.close:SET RST=NOTHING %>
			</TD> 		
			<TD nowrap align=right >組/部門<br>Bo phan</TD>
			<TD >
				<select name=g1  class=font9  onchange="submit()">
				<option value="" selected >--ALL---</option>
				<%
				SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
				'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
				SET RST = CONN.EXECUTE(SQL)
				'RESPONSE.WRITE SQL 
				WHILE NOT RST.EOF  
				%>
				<option value="<%=RST("SYS_TYPE")%>" <%if g1=RST("SYS_TYPE") then%>selected<%end if%> ><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
				<%
				RST.MOVENEXT
				WEND 
				%>
				</SELECT>
				<%rst.close: SET RST=NOTHING %>
			</td>		
		
			<td nowrap align=right >員工編號<br>So the</td>
			<td colspan=3>
				<input name=eid class=inputbox size=15 maxlength=5 value="<%=eid%>"> 
				
			</td>
		</TR> 	 
	</table> 
</td></tr></table> 	
<table width=470 align=center>
		<tr >
			<td align=right>
				<input type=button  name=btm class=button value="(Y)Confirm" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="(N)Cancel">	&nbsp;			 
				<input type=button  name=btm class=button value="(X)Close" onclick="window.close()">
			</td>
		</tr>	
	</table>
<table  border=0 cellspacing="1" cellpadding="1"  class="txt8">   
	<tr bgcolor="#dbdbdb" height=22>
		<td align="center" width=30 nowrap>STT</td>
		<td align="center" width=100 nowrap>部門單位</td>
		<td align="center" width=60 nowrap>工號</td>
		<td align="center" width=120 nowrap>姓名</td>
		<td align="center" width=80 nowrap >到職日</td>
		<td align="center">職務</td>
		<td align="center">特別獎金</td>
		<td align="center">USD<br>Rate:<%=rate%></td>
	</tr>
	<%while not rs.eof  
		x= x +1
		if x mod 2 = 0 then wkclr="#e4e4e4" else wkclr ="lightyellow"
		if rs("money1")="0" then money1="" else  money1 = rs("money1")
		if rs("money1")="0" then money1_usd="" else  money1_usd = round( cdbl(rs("money1"))/cdbl(rate) ,0)
	%>
	<tr bgcolor="<%=wkclr%>">
		<td align="center"><%=x%><input name="empid" value="<%=rs("empid")%>" type="hidden">
		<input name="country" value="<%=rs("country")%>" type="hidden"></td>
		<td align="left"><%=rs("gstr")%></td>
		<td align="center"><%=rs("empid")%></td>
		<td align="left"><%=rs("empnam_cn")%><br><%=rs("empnam_vn")%></td>
		<td align="center"><%=rs("nindat")%></td>
		<td align="left"><%=rs("jstr")%></td>
		<td align="left"><input class="inputbox" name="money1" value="<%=money1%>" style="text-align:right" size=10 ></td>
		<td align="left"><input class="inputbox" name="money1_USD" value="<%=money1_usd%>" style="text-align:right" size=10 ></td>
	</tr>
	<%rs.movenext
	wend 
	rs.close : set rs=nothing 
	%>
</table>
	<table width=470 align=center>
		<tr height=50>
			<td align=right>
				<input type=button  name=btm class=button value="(Y)Confirm" onclick="go()" onkeydown="go()">
				<input type=reset  name=btm class=button value="(N)Cancel">				 
				&nbsp;<input type=button  name=btm class=button value="(X)Close" onclick="window.close()">
			</td>
		</tr>	
	</table>
<script  type="text/javascript">
function gob(){
	var m = document.forms[0];
	var c1 = m.yymm.value ; 
	var c2 = m.country.value ; 
	var c3 = m.whsno.value ; 
	var c4 = m.groupid.value ; 
	var c5 = m.empid1.value ; 
	//alert (c1);
	window.open ("yece12.foreB.asp?yymm="+c1+"&country="+c2+"&whsno="+c3+"&g1="+c4+"&eid="+c5,"_blank","top=100,left=150,width=800,height=500,scrollbars=yes,resizable=yes" ) ;
}

function go(){
document.forms[0].action="yece12.foreB.asp?btn=GO";
document.forms[0].submit();
}
</script>	


</form>
</body>
</html>
 
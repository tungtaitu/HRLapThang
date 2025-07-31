<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%
Set conn = GetSQLServerConnection()	  
self="yebbb03"  


nowmonth = year(date())&right("00"&month(date()),2)  
if month(date())="01" then  
	calcmonth = year(date()-1)&"12"  	
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)   
end if 	   

if day(date())<=11 then 
	if month(date())="01" then  
		calcmonth = year(date()-1)&"12" 
	else
		calcmonth =  year(date())&right("00"&month(date())-1,2)   
	end if 	 	
else
	calcmonth = nowmonth 
end if   

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 20    'number of fields per record    

yymm=request("yymm")
yymm2=request("yymm2")
if request("yymm2")="" then yymm2=yymm
whsno=request("whsno")
country=request("country")
groupid=request("groupid")
empid1=request("empid1")
rpno=request("F_rpno")
rp_type = request("rp_type")
sortby = request("sortby")
if sortby="" then sortby="rpno desc"

if yymm="" and whsno="" and country="" and groupid="" and empid1="" and rpno="" then 
	sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
		"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112)='xxx' ) a "&_
		"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
		"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
		"where a.rpwhsno like '"& session("rpwhsno") &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
		"and a.empid like '"& empid1 &"%' and a.rp_type like '"&rp_type&"%' "&_
		"order by " & sortby 
else 
	if yymm<>""  then 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) between '"& yymm &"' and '"&yymm2&"' ) a "
	else 
		sql="select c.sys_value,b.whsno,b.empnam_cn,b.empnam_vn,b.nindat,b.gstr,b.zstr,b.groupid,b.zuno,b.job,b.jstr,a.* from "&_
				"(select convert(char(10),rp_dat, 111) as rp_date, * from emprepe where isnull(status,'')<>'D' and convert(char(6),rp_dat,112) like  '"& yymm &"%'  ) a  "
	end if  

	sql=sql&"left join ( select *from view_empfile ) b on b.empid = a.empid "&_
			"left join ( select *from basicCode ) c on c.func=case when a.rp_type='R' then 'goods' else case when a.rp_type='P' then 'bads' else '' end end  and c.sys_type = a.rp_func "&_
			"where a.rpwhsno like '"& WHSNO &"%' and  b.groupid like '"& groupid &"%' and b.country like '"&country&"%' "&_
			"and a.empid like '%"& empid1 &"%' and left(rpno,2) like '"& left(rpno,2) &"%'and a.rp_type like '"&rp_type&"%' "&_
			"order by  " & sortby 	
end if 
Set rs = Server.CreateObject("ADODB.Recordset") 
'response.write sql
'response.end 
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 1, 3 
	IF NOT RS.EOF THEN 	
		'PageRec = rs.RecordCount 
		rs.PageSize = PageRec 
		RecordInDB = rs.RecordCount 
		TotalPage = rs.PageCount  
		gTotalPage = TotalPage
	END IF 	 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	
	for i = 1 to TotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("rpno")
				tmpRec(i, j, 2) = rs("rp_date")
				tmpRec(i, j, 3) = rs("empid")
				tmpRec(i, j, 4) = rs("rp_type")
				tmpRec(i, j, 5) = rs("rp_func")
				tmpRec(i, j, 6) = rs("rp_method")
				tmpRec(i, j, 7) = rs("empnam_cn")
				tmpRec(i, j, 8) = rs("empnam_vn")
				tmpRec(i, j, 9) = rs("nindat")
				tmpRec(i, j, 10) = rs("whsno")
				tmpRec(i, j, 11) = rs("groupid")
				tmpRec(i, j, 12) = rs("zuno")
				tmpRec(i, j, 13) = rs("gstr")
				tmpRec(i, j, 14) = rs("job")
				tmpRec(i, j, 15) = rs("jstr")
				if rs("rp_type")="R" then
					tmpRec(i, j, 16)=rs("rp_type")&" 獎勵"
				elseif rs("rp_type")="P" then 
					tmpRec(i, j, 16)=rs("rp_type")&" 懲罰"
				else
					tmpRec(i, j, 16)=""
				end if 	
				tmpRec(i, j, 17) = tmpRec(i, j, 5) & " " & rs("sys_value")
				tmpRec(i, j, 18)=RS("FILENO")
				if rs("rp_type")="P" then 
					tmpRec(i, j, 19)="black"
				else
					tmpRec(i, j, 19)="blue"
				end if 	
				tmpRec(i, j, 20) = left(rs("rpmemo"),15)&"..."
				rs.MoveNext 
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
	Session("yebbb03B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("yebbb03B")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(TotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = TotalPage
	     Case Else
		      CurrentPage = 1
	end Select	
end if	
	

%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">   
</head>  
<body onkeydown="enterto()" onload='f()' >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
<INPUT TYPE=hidden NAME=sortby VALUE="<%=sortby%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
		<tr>
			<td>
				<table  class="txt" cellspacing="3" cellpadding="3"> 		 
					<TR>
						<TD nowrap align=right height=30 >廠別<BR><font class=txt8>Xuong</font> </TD>
						<TD > 
							<select name=WHSNO   style="width:150px">
								<option value="">全部ALL</option>
								<%
								if session("rights")=0 then 
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
								else
									SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
								end if 	
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=session("mywhsno") then%>selected<%end if%>><%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								SET RST=NOTHING
								%>
							</SELECT>
						</TD>  			 
						<TD nowrap align=right height=30 >統計<br>年月</TD>
						<TD colspan=3 nowrap>
							<input type="text" style="width:100px" name=yymm  size=8  value="<%=yymm%>"  maxlength=6 >~
							<input type="text" style="width:100px" name=yymm2  size=8  value="<%=yymm2%>" maxlength=6 >&nbsp;&nbsp;(yyyymm)
						</td>
						<td>	
							<input type=button name=btn value="(N)資料新增" class="btn btn-sm btn-danger" onclick=gonew() >
						</td>
					</tr>
					<tr>
						<TD nowrap align=right height=30 >國籍<BR><font class=txt8>Quoc Tich</font></TD>
						<TD >
							<select name=country  >
								<option value="">----</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='country'  ORDER BY SYS_type desc  "
								SET RST = CONN.EXECUTE(SQL)
								WHILE NOT RST.EOF  
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=country then%>selected<%end if%>><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
								<%
								RST.MOVENEXT
								WEND 
								SET RST=NOTHING
								%>
							</SELECT>
						</TD>					
						<TD nowrap align=right >部門<BR><font class=txt8>Don vi</font></TD>
						<TD >
							<select name=GROUPID style="width:150px">
							<option value="" selected >----</option>
							<%
							SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
							'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
							SET RST = CONN.EXECUTE(SQL)
							'RESPONSE.WRITE SQL 
							WHILE NOT RST.EOF  
							%>
							<option value="<%=RST("SYS_TYPE")%>" <%if RST("SYS_TYPE")=groupid then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>				 
							<%
							RST.MOVENEXT
							WEND 
							%>
							</SELECT>
							<%rst.close
							SET RST=NOTHING 
							conn.close 
							set conn=nothing
							%>
						</td>
						<td nowrap align=right >工號<BR><font class=txt8>So The</font></td>
						<td>
							<input type="text" style="width:100px" name=empid1  size=10 maxlength=5   value="<%=empid1%>">
						</td>
						<td>	
							<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(S)查詢K.tra" onclick="go()" onkeydown="go()">
						</td>
					</TR>  
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid">
					<tr class="header">
						<Td  nowrap  >STT</td>
						<Td  nowrap  >獎懲<br>thuong phat</td>
						<Td  nowrap  onclick="dchg(1)" style="cursor:pointer">編號<br>so<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap  onclick="dchg(2)" style="cursor:pointer">事件日期<br>Ngay<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap  onclick="dchg(3)" style="cursor:pointer" >工號<br>so the<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap  >姓名<br>ho ten<br></td>
						<Td  nowrap  onclick="dchg(4)" style='cursor:pointer'>部門<br>bo phan<br><img src="../picture/soryby.gif"></td>		
						<Td  nowrap  onclick="dchg(5)" style='cursor:pointer'>方式<br>phuong thuc<br><img src="../picture/soryby.gif"></td>
						<Td  nowrap   >內容<br>thuyet minh</td>
						<Td  nowrap  >處理說明<br>phuong thuc xu ly </td>		
						<Td  nowrap  >文件編號<br>so van kien</td>		
					</tr>
					<%for x = 1 to pagerec
						if x mod 2 = 0 then 
							wkcolor=""
						else
							wkcolor="PapayaWhip"
						end if 	
					%>
						<Tr style="cursor:pointer"  onclick='chkdata(<%=x-1%>)' class="txt8">
							<Td align=center><%=(currentpage-1)*pagerec+x%></td>
							<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 16)%></font></td>
							<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 4)%><%=tmpRec(CurrentPage, x, 1)%></font></td>
							<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 2)%></font></td>
							<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 3)%></font></td>
							<Td><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 7)%><BR><%=tmpRec(CurrentPage, x, 8)%></font></td>
							<Td align=center><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 13)%></font></td>			
							<Td  width=70  align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 17)%></font></td>
							<Td width=120><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 20)%></font></td>
							<Td width=120 ><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 6)%></font>
								<input type=hidden name=empid value="<%=tmpRec(CurrentPage, x, 3)%>">
								<input type=hidden name=rpno value="<%=tmpRec(CurrentPage, x, 1)%>">
								<input type=hidden name=rptype value="<%=tmpRec(CurrentPage, x, 4)%>">
							</td>
							<Td align=left><font color="<%=tmpRec(CurrentPage, x, 19)%>"><%=tmpRec(CurrentPage, x, 18)%></font></td>
						</tr>
					<%next%>
				</table>
			</td>
		</tr>
		<tr align="center">
			<td>
				<input type=hidden name=empid>
				<input type=hidden name=rpno>
				<input type=hidden name=rptype>
				<table class=txt8 cellpadding=3 cellspacing=3>
					<tr><Td height=25 align=center>page : <%=currentpage%>/<%=totalpage%>, recordCount:<%=recordIndB%></td></tr>
					<Tr>
						<td align="CENTER" height=40 >    
						<% If CurrentPage > 1 Then %>
							<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
						<% Else %>
							<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
						<% End If %>
						<% If cint(CurrentPage) < cint(TotalPage) Then %>
							<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
						<% Else %>
							<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
							<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
						<% End If %>
						</td>
						<td width="200px" align="center">
							<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="轉 EXCEL" onclick=goexcel()  >
						</td>									
					</tr>
				</table>
			</td>
		</tr>
	</table>
			
	
</form>
</body>
</html>
<script language=javascript> 
function enterto(){
	if(window.event.keyCode == 13) window.event.keyCode =9;
}

function dchg(a) {
	switch (a){
		case 1 :
			<%=self%>.sortby.value="rpno desc";
			break;
		case 2 :
			<%=self%>.sortby.value="rp_dat, rpno ";
			break;
		case 3 :
			<%=self%>.sortby.value="a.empid, rp_dat";
			break;
		case 4 :
			<%=self%>.sortby.value="b.groupid, a.empid, a.rp_dat";
			break;
		case 5 :
			<%=self%>.sortby.value="rp_func, rpno ";
			break;
	} 	
	<%=self%>.TotalPage.value="0";
 	<%=self%>.action="<%=self%>.Fore.asp";
 	<%=self%>.submit();													
}

function f(){
	<%=self%>.yymm.focus();
}

function gonew() {   
	open("<%=self%>.new.asp", "_self");
} 

function go2(a)
{
	if(a==1)
	{ 
		if(<%=self%>.yymm.value !="")
		{ 
			<%=self%>.TotalPage.value="0";
		 	<%=self%>.action="<%=self%>.Fore.asp";
		 	<%=self%>.submit() ;
		}
	}
	else if(a==2)
	{ 
		if(<%=self%>.empid1.value !="")
		{
			<%=self%>.TotalPage.value="0";
		 	<%=self%>.action="<%=self%>.Fore.asp";
		 	<%=self%>.submit();
		} 	
	}  
}  
	
function go() {
	<%=self%>.TotalPage.value="0";
 	<%=self%>.action="<%=self%>.Fore.asp";
 	<%=self%>.submit() ;
} 

function chkdata(index){
	 rpno= <%=self%>.rpno[index].value;
	 empid = <%=self%>.empid[index].value;
	 rptype= <%=self%>.rptype[index].value;
	 open("<%=self%>.foregnd.asp?rpno="+rpno +"&empid="+empid+"&rptype="+rptype, "_self" );
}  
	
function goexcel(){	 
	<%=self%>.action="<%=self%>.toexcel.asp";
	<%=self%>.target="Back";
	<%=self%>.submit();
	parent.best.cols="100%,0%";
}
</script> 
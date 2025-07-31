<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

self="YECE1303"  
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

Set conn = GetSQLServerConnection()

years = request("years")
ct = request("ct")
whsno = request("whsno")

if instr(conn,"168.")>0 then 
	w1="LA" 
elseif instr(conn,"169.")>0 then 	
	w1="DN" 
elseif instr(conn,"47.")>0 then 	
	w1="BC" 
else
	w1=""
end if  
w1=session("mywhsno")

flag=request("flag")
g1= request("groupid")
eid= request("eid") 
c1 = request("c1")
khud=request("khud") 
gTotalPage = 1
PageRec = 1    'number of records per page
TableRec = 15    'number of fields per record    

'response.write "xxx="& khud 

enddat =request("years")&"/12/31" 
if request("years")="" then enddat =year(date())&"/12/31" 
Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array  
if flag="S" then 
	
	sql="select isnull(c.years,'"&years&"') as years, a.empid, a.empnam_cn, a.empnam_vn, a.country, convert(char(10),a.indat,111) as indate, b.whsno, b.groupid , "&_ 
			"isnull(convert(char(10),outdat,111),'') outdate, ix.sys_value as gstr, datediff(m,a.indat,'"&enddat&"')/1.0 as nz  , isnull(c.fensu,'') fensu, isnull(c.kj,'') kj "&_
		  "from "&_
			"( select *  from empfile  where  ( isnull(outdat,'')='' or convert(char(10),outdat,111)>='"& enddat &"' )  and isnull(status,'')<>'D' ) a "&_
			"left join (select *from bempg where  yymm=convert(char(6),getdate(),112) ) b on b.empid = a.empid "&_
			"left join (Select * from basiccode where func='groupid' ) ix on ix.sys_type = b.groupid "&_
			"left join (select * from empnzkh where years='"&years&"' and khud='"&khud&"') c on c.empid = a.empid  "&_
			"where b.whsno='"&whsno &"' and b.groupid like'"&g1&"%' and a.empid like '"&eid&"%' and a.country like '"&c1&"%' "&_
			"order by b.whsno, a.country, b.groupid, a.empid  "
	'response.write sql
	'response.end 
	Set rs = Server.CreateObject("ADODB.Recordset") 
	rs.open sql, conn, 3, 3 
	
	CurrentPage = 1	 
	
	if not rs.eof then 		
		pagerec= rs.recordcount 				
		rs.pagesize = pagerec 
		recordindb = rs.recordcount 
		totalpage = rs.pagecount  
		gtotalpage = totalpage
		'whsno = rs("whsno")
	end if 	
	
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
	for i = 1 to gTotalPage 
		for j = 1 to PageRec
			if not rs.EOF then 			
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("whsno"))
				tmpRec(i, j, 2) = trim(rs("years"))
				tmpRec(i, j, 3) = trim(rs("country"))
				tmpRec(i, j, 4) = rs("empid")
				tmpRec(i, j, 5) = rs("indate")
				tmpRec(i, j, 6) = rs("groupid")
				tmpRec(i, j, 7) = rs("gstr")				
				tmpRec(i, j, 8) = rs("nz")
				tmpRec(i, j, 9) = rs("fensu")
				tmpRec(i, j, 10) = rs("kj")
				tmpRec(i, j, 11) = rs("empnam_cn")
				tmpRec(i, j, 12) = rs("empnam_vn")
				tmpRec(i, j, 13) = rs("outdate")
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
	session("yece1303b")=tmprec
end if 


%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()"  >
<form name="<%=self%>" method="post"  ENCTYPE="multipart/form-data"  >

<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<TD nowrap align=right height=30 ><font color="blue">檔案<br>File</font></TD>			
									<Td colspan=5 nowrap>
										<INPUT TYPE="FILE" NAME="FILE1" style="width:200px">
										<input type="button" name=btn value="(Y)上傳upload" class="btn btn-sm btn-outline-secondary" onclick="go()">										
									</td>
								</tr>				
								<tr>			
									<TD nowrap align=right>廠別<br>Xuong</TD>
									<TD> 
										<select name="WHSNO" style="width:120px">
											<option value="">----</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if w1=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD> 	
									<TD nowrap  align=right height=30 >年度<br>Nam</TD>			
									<TD nowrap>
										<input  type="text" style="width:100px" name="years" maxlength=4 value="<%=years%>" >
										<select name="khud" style="width:50px">
											<option value="" <%if khud="" then%>selected<%end if%> > -----</option>
											<option value="0" <%if khud="0" then%>selected<%end if%> >上</option>
											<option value="1" <%if khud="1" then%>selected<%end if%> >下</option>
										</select>												
									</td>	 
								</TR>		 
</form>
<form name="<%=self%>B" method="post">	
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 		
								<tr >	
									<TD nowrap  align=right height=30 >國籍<br>Quoc tich</TD>			
									<TD > 
										<select name="c1"   style="width:120px"  >
											<option value="">----</option>
											<%'SQL="SELECT * FROM BASICCODE WHERE FUNC='country' ORDER BY SYS_TYPE "
											if session("rights")<=0  then 				
												sql="select *from basiccode where func='country' order by sys_type" 
											else	
												sql="select *from basiccode where func='country' and sys_type in ('VN','TA') order by sys_type" 
											end if 						
											SET RST = CONN.EXECUTE(SQL)
											if c1="" and w1<>"LA" then c1="VN"
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if c1=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>		 
									<TD nowrap  align=right height=30 >部門<br>bo phan</TD>			
									<TD > 
										<select name="groupid"  style="width:120px"  >
											<option value="">----</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='groupid' and sys_type<>'AAA' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if g1=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>			
									<TD nowrap   align=right   >工號<br>so the</TD>
									<TD nowrap><input  type="text" style="width:100px"  name="eid"  maxlength=5 value="<%=eid%>"  ></TD> 
									<td nowrap>
										<table border=0 class="txt" cellpadding=3 cellspacing=3>
											<tr>
												<td><input name=btn value="(S)查詢" type="button" class="btn btn-sm btn-outline-secondary" onclick="gos()"></td>
												<td><input name=btn value="(C)Cancel" type="button" class="btn btn-sm btn-outline-secondary" onclick="goc()"></td>
												<td><input name=btn value="To Excel" type="button" class="btn btn-sm btn-outline-secondary" onclick="goexcel()"></td>
											</tr>
										</table>
									</td>
								</TR>				
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrtid" width="98%">
								<tr bgcolor=#e4e4e4 class="txt8">
									<td align="center">STT</td>
									<td align="center">廠別<br>xuong</td>
									<td align="center">年度<br>nam</td>
									<td align="center">國籍<br>quoc<br>tich</td>
									<td align="center">單位<br>bo phan</td>
									<td align="center">工號<br>so the</td>
									<td align="center">姓名<br>ho ten</td>
									<td align="center">到職日(NVX)<br>離職日(NTV)</td>
									<td align="center">年資<br>so thang<br>lam viec</td>		
									<td align="center">分數<br>fensu</td>
									<td align="center">考績<br>grade</td>
								</tr>
								<%for x = 1 to pagerec 
								 if x mod 2 = 0 then 
									wkclr="#ffffff"
								 else	
									wkclr="#ffffcc"
								 end if
								%>
								<tr bgcolor="<%=wkclr%>">
									<td align="center"><%=x%></td>
									<td align="center"><%=tmprec(1,x,1)%></td>
									<td align="center"><%=tmprec(1,x,2)%></td>
									<td align="center"><%=tmprec(1,x,3)%></td>
									<td align="center"><%=tmprec(1,x,7)%></td>
									<td align="center"><%=tmprec(1,x,4)%></td>
									<td><%=tmprec(1,x,11)%><br><%=tmprec(1,x,12)%></td>
									<td align="center"><%=tmprec(1,x,5)%><br><font color="red"><%=tmprec(1,x,13)%></font></td>
									<td align="center"><%=tmprec(1,x,8)%></td>
									<td align="center">
										<% if tmprec(1,x,5)>=years&"/07/02" then %> 0 
											<input type="hidden" name="fensu" value="0" class="inputbox8" size=2 style="text-align:center"  >
										<%else%>
											<input type="text" name="fensu" value="<%=tmprec(1,x,9)%>" class="inputbox8" style="width:100%;text-align:center" onblur="fschg(<%=x-1%>)">
										<%end if%>	
									</td>
									<td align="center">
										<% if tmprec(1,x,5)>=years&"/07/02" then %> N
										<input type="hidden" name="grade" value="N" class="inputbox8" size=2 style="width:100%;text-align:center"  >
										<%else%>
										<input type="text" name="grade" value="<%=tmprec(1,x,10)%>" class="inputbox8" style="width:100%;text-align:center" onblur="kjchg(<%=x-1%>)" >
										<%end if%>
									</td>
								</tr>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<input name="fensu" type="hidden">
							<input name="grade" type="hidden">							
							<table  class="txt" cellpadding=3 cellspacing=3>
								<Tr>
									<td align="center">
										<input  type="button" name="btn" value="(Y)Confirm" class="btn btn-sm btn-danger" onclick="goupd()">
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</table>
			
</form>  

</body>
</html>
 
 
<script language=vbs> 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		 	
end function 

function goupd()
	if <%=self%>.whsno.value="" then 
		alert "請輸入廠別xin dnah lai xuong"
		<%=self%>.whsno.focus()
		exit function 
	end if 	
	if <%=self%>.years.value="" then 
		alert "請輸入年度xin dnah lai Nam"
		<%=self%>.years.focus()
		exit function 
	end if 
	if <%=self%>.khud.value="" then 
		alert "請輸入上下年度"
		<%=self%>.khud.focus()
		exit function 
	end if 
	if <%=self%>B.fensu(0).value="" then 
		alert "無資料No data!!"
		exit function 
	end if 
	w1=<%=self%>.whsno.value
	years=<%=self%>.years.value
	khud=<%=self%>.khud.value
	<%=self%>B.action="<%=self%>.insDB.asp?whsno="&w1 &"&years="& years &"&khud="& khud 
	<%=self%>B.target="_self"
	<%=self%>B.submit()
end function   

function goexcel()
	code1=<%=self%>.years.value
	code2=<%=self%>.whsno.value
	code3=<%=self%>B.groupid.value
	code4=<%=self%>B.eid.value
	code5=<%=self%>B.c1.value
	code6=<%=self%>.khud.value
	<%=self%>.action = "<%=self%>.toexcel.asp?flag=S&years="&code1 &"&whsno="& code2 &"&groupid="& code3 &"&eid="& code4 &"&country="& code5 &"&khud="& code6
	<%=self%>.target="Back"
	<%=self%>.submit()
	'parent.best.cols="50%,50%"
end function  

function  fschg(index)
	if <%=self%>b.fensu(index).value<>"" then 
		if isnumeric(<%=self%>b.fensu(index).value)=false then 
			alert "請輸入數字!!xin danh lai so"
			<%=self%>b.fensu(index).value="0"
			<%=self%>b.grade(index).value="N"
			<%=self%>b.fensu(index).select()
		else
			fs=cdbl(<%=self%>b.fensu(index).value)
			if fs>=90 then 
				<%=self%>b.grade(index).value="優" 
			elseif fs>=85 then 	
				<%=self%>b.grade(index).value="良" 
			elseif fs>=80 then 	
				<%=self%>b.grade(index).value="甲" 
			elseif fs>=70 then 	
				<%=self%>b.grade(index).value="乙" 
			elseif fs<70 and fs>0 then 	
				<%=self%>b.grade(index).value="丙" 
			elseif fs=0 then 
				<%=self%>b.grade(index).value="N" 
			else
				<%=self%>b.grade(index).value=""
			end if 
		end if 
	end if 
end function 

function kjchg(index)
end function 

function f()
 	<%=self%>.years.focus() 
end function   
function clr()
	open "<%=SELF%>.asp" , "_self"
end function  

function goc()
	open "<%=self%>.asp" , "_self"
end function 
  
function go() 
	 if <%=self%>.years.value="" then 
		alert "請輸入年度xin danh lai nam"
		<%=self%>.years.focus()
		exit function 
	 end if 	
	 if <%=self%>.whsno.value="" then 
		alert "請輸入廠別xin danh lai xuong"
		<%=self%>.whsno.focus()
		exit function 
	 end if  

	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.target="_self"
 	<%=self%>.submit() 
end function  

function gos()
	code1=<%=self%>.years.value
	code2=<%=self%>.whsno.value
	code3=<%=self%>B.groupid.value
	code4=<%=self%>B.eid.value
	code5=<%=self%>B.c1.value
	code6=<%=self%>.khud.value
	<%=self%>.action = "<%=self%>.fore.asp?flag=S&years="&code1 &"&whsno="& code2 &"&groupid="& code3 &"&eid="& code4 &"&c1="& code5 &"&khud="& code6
	<%=self%>.target="_self"
	<%=self%>.submit()
end function 
 
</script> 
<%@Language=VBScript codepage=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" --> 
<!-- #include file="../../ADOINC.inc" -->
<!--#include file="../../include/checkpower.asp"-->
<!--#include file="../../include/sideinfolev2.inc"-->
<%
SELF = "EMPBASICB"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")  
Set rst = Server.CreateObject("ADODB.Recordset")  

gTotalPage = 1
PageRec = 32    'number of records per page
TableRec = 10    'number of fields per record  

'Response.Write nowmonth &"<BR>"
'Response.Write calcmonth &"<BR>"      
'Response.End 
'a=4.35689 
'b = - Int(-a)
'response.write b &<BR>" 

yymm = request("yymm")   
if trim(request("yymm2"))<>"" then yymm = trim(request("yymm2"))


if yymm<>"" then 
	chkym=left(yymm,4)&"/"&right(yymm,2)
	cDatestr=CDate(LEFT(YYMM,4)&"/"&RIGHT(YYMM,2)&"/01") 
	days = DAY(cDatestr+(32-DAY(cDatestr))-DAY(cDatestr+(32-DAY(cDatestr))))   '一個月有幾天
	
	sql="select * from YDBMCALE where convert(char(6),dat,112)='"& yymm &"' order by dat "
	rst.open sql, conn, 3, 3 
	if rst.eof then 
		for k = 1 to days 
			insdate = chkym & "/"&right("00"&k,2)
			if weekday(insdate)=1 then 
				stas="H2" 
			else
				stas="H1"
			end if 		
			sql =" insert into YDBMCALE (dat, status ) values ('"& insdate &"' , '"& stas &"' ) "
			'response.write sql &"<BR>"
			conn.execute(sql)
			'response.write sql 
		next 			
	end if 
	set rst=nothing 
	'----------------------------------------------------------------------------------------------
	if request("TotalPage") = "" or request("TotalPage") = "0" then 
		CurrentPage = 1	
		sqlstr = "select * from YDBMCALE where convert(char(6),dat,112)='"& yymm &"' order by dat "
		rs.Open SQLstr, conn, 3, 3 
		IF NOT RS.EOF THEN 
			rs.PageSize = PageRec 
			RecordInDB = rs.RecordCount 
			TotalPage = rs.PageCount  
			gTotalPage = TotalPage
		END IF 	 
	
		Redim tmpRec(TotalPage, PageRec, TableRec)   'Array
		
		for i = 1 to TotalPage 
			for j = 1 to PageRec
				if not rs.EOF then 					
					tmpRec(i, j, 0) = "no"
					tmpRec(i, j, 1) = year(trim(rs("dat")))&"/"&right("00"&month(rs("dat")),2)&"/"&right("00"&day(rs("dat")),2)
					tmpRec(i, j, 2) = trim(rs("status"))										
					tmpRec(i, j, 3)= mid("日一二三四五六",weekday(tmpRec(i, j, 1)) , 1 )	
					
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
		'Session("EMPBASICB") = tmpRec	 
	end if 	 
else
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array 
end if 

  
 
%>

<html>

<head>

<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">


<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css"> 

<link rel="stylesheet" type="text/css" href="../../template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../../template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="../../template/css/mis.css">
<link rel="stylesheet" type="text/css" href="../../template/datepicker/datepicker.css">

  
</head>   
<body  topmargin="0" leftmargin="0"  marginwidth="0" marginheight="0" >
<form name="<%=self%>" method="post"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="60%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt">
								<tr>
									<td>設定年月<br>Thiết lập ngày：</td>
									<td>
										<select name=yymm style="vertical-align:middle;width:100px"  onchange="datachg()">
											<option value=""></option>
											<%for z = 1 to 24 
											  if z > 12 then 
													zz = z mod 12 
													if zz = 0 then zz = 12  
													n_ym = year(date())+1
												else
													zz = z 
													n_ym = year(date())
												end if 	
												yymmvalue = n_ym&right("00"&zz,2)
											%>
												<option value="<%=yymmvalue%>" <%if yymmvalue=yymm then %>selected<%end if%>><%=yymmvalue%></option>
											<%next%>	
										</select>  				
									</td> 			
									<td><input type="text" style="width:50px" readonly  name=days value="<%=days%>" >  </td>  
									<%if session("netuser")="PELIN" then %>
									<td><input type="text" style="width:100px"  name=yymm2 value=""   onblur="datachg()" ></td>  
									<%else%>
										<input  name=yymm2 value="" size=7  type="hidden">
									<%end if %>
								</tr>		
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="txt" width="100%">
								<tr bgcolor=#cccccc height=30>
									<td width=90 align=center height=35 >日期<br>Ngày</td>
									<td width=40 align=center>星期<br>Thứ</td>			
									<td width=40 align=center>假日<br>Ngày nghỉ</td>
									<td width=40 align=center>國定<br>假日<br>Ngày lễ</td>
									<td width=5 align=center bgcolor=#ffffff>
									<td width=90 align=center>日期<br>Ngày</td>
									<td width=40 align=center>星期<br>Thứ</td>			
									<td width=40 align=center>假日<br>Ngày nghỉ</td>
									<td width=40 align=center >國定<br>假日<br>Ngày lễ</td>
									<td width=5 align=center bgcolor=#ffffff>
								</tr>
								<%for x = 1 to days 				
								%>
								<%if x mod 2 = 1 then %><tr bgcolor="Beige"><%end if %>
									<td align=center>

										<%if weekday(tmprec(1,x,1))=1 then %>
											<font color=red><%=tmprec(1,x,1)%></font>
										<%else%>	
											<font color=black><%=tmprec(1,x,1)%></font>
										<%end if%>

										<input type=hidden name=dat  size=2 value=<%=tmprec(1,x,1)%> class=inputbox>
										<input type=hidden name=b_sts  size=2 value=<%=tmprec(1,x,2)%> class=inputbox>
										<input type=hidden name=status  size=2 value=<%=tmprec(1,x,2)%>  class=inputbox >
									</td>
									<td align=center><%=tmprec(1,x,3)%></td>					 
									<td align=center>
										<%if tmprec(1,x,2)="H2" then  %>
											<input type=checkbox name=T2 checked  onclick="t2chg(<%=x-1%>)" >
										<%else%>	
											<input type=checkbox name=T2 onclick="t2chg(<%=x-1%>)" >
										<%end if%>	
									</td>
									<td align=center >
										<%if tmprec(1,x,2)="H3" then  %>
											<input type=checkbox name=T3 checked  onclick="t3chg(<%=x-1%>)" value="t3" >
										<%else%>	
											<input type=checkbox name=T3  onclick="t3chg(<%=x-1%>)" value="t3">
										<%end if%>	
									</td>
									<td width=5 bgcolor=#ffffff >	
								<%if x mod 2 = 0 then %></tr><%end if %>
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="txt" width="100%">
								<tr><td>&nbsp;</td></tr>
								<tr>
									<TD align=center>				 
									<%if UCASE(session("mode"))="W" then%>
									<input type="button" name=send value="確　　認 XÁC NHẬN"  class="btn btn-sm btn-danger" onclick="go()" onkeydown="go()" > 
									<input type=RESET name=send value="取　　消 HỦY BỎ"  class="btn btn-sm btn-outline-secondary">		
									<%end if%>
									</TD>
								</TR>
							</table>
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
function BACKMAIN() 	
	open "../main.asp" , "_self"
end function 

function datachg()
	if <%=self%>.yymm.value<>"" or <%=self%>.yymm2.value<>"" then 
		<%=self%>.action = "empbasicB.fore.asp"
		<%=self%>.submit()
	end if 	
end function  

function t2chg(index)
	if <%=self%>.t2(index).checked=true then 
		if <%=self%>.t3(index).checked=true then 
			<%=self%>.t3(index).checked=false 
		end if 
		<%=self%>.status(index).value="H2"
	else
		' <%=self%>.status(index).value=<%=self%>.b_sts(index).value 
		' if <%=self%>.b_sts(index).value="H3" then 
			' <%=self%>.t3(index).checked=true 
		' end if 	
		<%=self%>.status(index).value="H1"
	end if 	 
end function 

function t3chg(index)
	if <%=self%>.t3(index).checked=true then 
		if <%=self%>.t2(index).checked=true then 
			<%=self%>.t2(index).checked=false 
		end if 
		<%=self%>.status(index).value="H3"
	else
		' <%=self%>.status(index).value=<%=self%>.b_sts(index).value 
		' if <%=self%>.b_sts(index).value="H2" then 
			' <%=self%>.t2(index).checked=true 
		' end if 	
		<%=self%>.status(index).value="H1"
	end if 	 
end function 

function go()
	<%=self%>.action = "empbasicB.upd.asp"
	<%=self%>.submit()
end function 
 
 
	
</script>



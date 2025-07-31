<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
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
'response.write tjnum 
'response.end 

sql="exec SP_empNJTJ_N '"& yymm &"', '"& country &"','"& whsno &"','"& groupid &"','"& empid1 &"','"& sortby &"' "  
'response.write sql 
'response.end 
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
	'RESPONSE.END 
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
				tmpRec(i, j, 13)=RS("tx_hrbh")
				tmpRec(i, j, 14)=RS("txdaysbh")
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
				tmpRec(i, j, 27)=rs("bhdate")
				
				for kk = 1 to tjnum 
					tmpRec(i, j, 27+kk) = rs(4+kk)  '' 每月特休時數
					tmpRec(i, j, 27+tjnum+kk) = rs(4+kk).name    '欄位名稱 T103H ( 第2碼表示年  ,ex : 2011  就是1 , 2012 = 2 , 3.4碼為月 ) 
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
	
	Session("YEEE03B") = tmpRec	
	
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
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function  

function f()
	'<%=self%>.QUERYX.focus()	
end function   

function gos()
	<%=SELF%>.totalpage.VALUE=""
	<%=self%>.action="<%=self%>.foregnd.asp"
	<%=self%>.target="Fore"	
	<%=self%>.submit
end function 


-->
</SCRIPT>  
</head>   
<body  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=SELF%>.foregnd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 
<INPUT NAME=yymm VALUE="<%=yymm%>" TYPE=HIDDEN  > 
<INPUT NAME=whsno VALUE="<%=whsno%>" TYPE=HIDDEN  > 
<INPUT NAME=country VALUE="<%=country%>" TYPE=HIDDEN  > 

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="98%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt">
								<Tr>
									<td class="txt" align="right">統計年度：</td>
									<td colspan=3  class="txt" ><%=yymm%> ( <%=td1%> ~ <%=td2%> )</td>
								</tr>
								<tr>	
									<TD nowrap align=right >部門：</TD>
									<TD>
										<select name=GROUPID  onchange="gos()" style="width:150px">
											<option value="" selected >全部 </option>
											<%
											SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
											'SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID'  ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											'RESPONSE.WRITE SQL
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if GROUPID=RST("SYS_TYPE") then%>selected<%end if%> ><%=RST("SYS_TYPE")%> <%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</td>
									<td nowrap align=right >排序：</td>
									<td nowrap>
										<select  name=showby onchange="gos()" style="width:150px">					
											<option value="A" <%if sortby="A" then%>selected<%end if%>>A.依部門/工號</option>
											<option value="B" <%if sortby="B" then%>selected<%end if%>>B.依工號</option>			
											<option value="" <%if sortby="" then%>selected<%end if%>>ALL</option>
										</select>
									</td>
									<td nowrap align=right >工號：</td>
									<td>
										<input type="text" style="width:100px" name=empid1 maxlength=5 value="<%=empid1%>" >
									</td>
									<td>									
										<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="(S)查詢" onclick="gos()" onkeydown="gos()">
									</td>
									<td>
										<input type=button  name=btn class="btn btn-sm btn-outline-secondary" value="save To Excel" onclick=goexcel() style='background-color:#ffccff' >
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table id="myTableGrid" width="98%">
								<tr BGCOLOR="LightGrey" class="txt9">		
									<TD nowrap align=center >STT</TD> 		
									<TD  nowrap align=center >工號<br>so the</TD> 		
									<TD nowrap align=center >姓名<br>ho ten</TD>
									<TD align=center  nowrap >到職日<br>NVX</TD>
									<TD align=center   nowrap>職務<br>chu vu</TD>		
									<TD align=center   nowrap >單位<br>Don vi</TD> 			
									<%for k1=1 to tjnum %>	
										<TD align=center  nowrap ><%=left(tmprec(1,1,27+tjnum+k1),4)%><br>H</TD>
									<%next%>
									<TD align=center >調整<br>Điều chỉnh</TD>
									<TD align=center  nowrap >年假(天)<br>Phép năm(ngày)</TD>
									<TD align=center  nowrap >年假(時數))<br>Phép năm(giờ)</TD>
									<TD align=center  nowrap >剩餘年假(時數)<br>已修時數<br>Phép năm còn lại(giờ) <br>Đã nghỉ</TD>
								</tr>
								 
								<%for x = 1 to PageRec
									IF x MOD 2 = 0 THEN 
										WKCOLOR="LavenderBlush"
									ELSE
										WKCOLOR="lightyellow"
									END IF 	 
									'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
								%>
								<TR BGCOLOR=<%=WKCOLOR%> class="txt9"> 	
									<TD align=center ><%=x%></td>
									<TD align=center ><a href="#" onclick="insMemo(<%=x-1%>)"><%=tmpRec(CurrentPage,x,1)%></a></td>
									<TD nowrap>
										<a href="vbscript:insMemo(<%=x-1%>)">
										<%=tmprec(1,x,2)%><br><%=left(tmprec(1,x,3),22)%>			
										</a>
									</td>
									<TD nowrap><%=tmprec(1,x,5)%>
									<br><%=tmprec(1,x,27)%>
									<br><font color="red"><%=tmprec(1,x,6)%></font>
									</td>
									<TD nowrap><%=left(tmprec(1,x,11),8)%></td>
									<TD align=left ><%=tmprec(1,x,9)%><br><%=tmprec(1,x,10)%></td>		
									<%for y = 1 to tjnum %>
											<TD align=center ><%if tmprec(1,x,27+y)<>"0" then %>
												<%=tmprec(1,x,27+y)%><%end if%>
											</td>
									<%next%> 		
									<TD align=center valign="top" >
										<%
										if cdbl(tmprec(1,x,25))=0 then 
											tz_txd="" 
											atz_txd = tmprec(1,x,14) 
											atz_txh = tmprec(1,x,13)																						
											nowtxH = cdbl(tmprec(1,x,13))- cdbl(tmprec(1,x,12))
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
										<input type="text" class="inputbox8" name="njtz" style="width:40px" onblur="njtzchg(<%=x-1%>)" value="<%=tz_txd%>">
									</td>
									<TD align="right"  nowrap="nowrap" valign="top">	
										
									<input type="text" class="readonly8" readonly name="txdays" value="<%=atz_txd%>" style="width:40px" >
									<br><font color="#999999"><%=round(tmprec(1,x,14),1)%></font> 
									<input  class="inputbox8" name="old_txdays" value="<%=tmprec(1,x,14)%>" type="hidden">
									<input  class="inputbox8" name="allJiaE" value="<%=tmprec(1,x,12)%>" type="hidden">
									</td>
									<TD align=center valign="top"><input type="text" class="readonly8" name="tx_hr" value="<%=atz_txh%>" style="width:40px" readonly ></td>		
									<TD align=right nowrap="nowrap">
										<input type="text" class="readonly8" name="nowtx" value="<%=nowtxH%>" style="width:40px;<%if nowtxH <0 then %> background-color:red <% end if%> " readonly >
										<br><font color="#999999"><%=round(tmprec(1,x,12),1)%></font>
									</td>	
									<%if session("rights")<="2" then %>
													
										<% if cdbl(atz_txh)>0 then%>
										<input type="hidden" class="readonly8" name="njAmt" value="<%=nj_amt%>" style="width:40px"  readonly> 
										<%else%>
										<input type="hidden" class="readonly8" name="njAmt" value="0" style="width:40px"  >
										<%end if%>
										
									<%else%>
										<input class="inputbox8" name="njAmt" size=5 value="0" type="hidden">			
									<%end if%>									
									<input class="inputbox8" name="hh_money" size=5 value="<%=formatnumber(tmprec(1,x,20),0)%>" type="hidden">
								</TR>
								<%next%> 	 
								<input class="inputbox8" name="njAmt" size=5 value="0" type="hidden">
								<input class="inputbox8" name="hh_money" size=5 value="0" type="hidden">
								<input class="inputbox8" name="njtz" size=5 value="0" type="hidden">
								<input class="inputbox8" name="txdays" size=5 value="0" type="hidden">
								<input class="inputbox8" name="old_txdays" size=5 value="0" type="hidden">
								<input class="inputbox8" name="tx_hr" size=5 value="0" type="hidden">
								<input class="inputbox8" name="nowtx" size=5 value="0" type="hidden">
								<input class="inputbox8" name="allJiaE" size=5 value="0" type="hidden">	
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
								  <td align="CENTER" class="txt8">
									
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
									PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%>
									</td>
									<td align=right>
										<%'if session("rights")<="2" then %>	
										<input type="BUTTON" name="send" value="確　定" class="btn btn-sm btn-outline-secondary" onclick="GO()" >
										<input type="BUTTON" name="send" value="取　消" class="btn btn-sm btn-outline-secondary" ONCLICK="CLR()">	
										<%'end if%>
									</td>		
								</TR>
							</TABLE> 
						</td>
					</tr>					
				</table>
			</td>
		</tr>
	</table>

</form>

</body>
</html>
 
<script language=vbscript> 
function goexcel()
	'open "<%=self%>.toexcel.asp" , "Back" 
	parent.best.cols="100%,0%"
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	
end function 

function njtzchg(index)
	if <%=self%>.njtz(index).value<>"" then 
		if isnumeric(<%=self%>.njtz(index).value)=false then 
			alert "請輸入數值!! xin danh lai [so]"
			<%=self%>.njtz(index).value="" 
			exit function 
		else			
			<%=self%>.txdays(index).value = cdbl(<%=self%>.old_txdays(index).value)+cdbl(<%=self%>.njtz(index).value)
			<%=self%>.tx_hr(index).value = cdbl(<%=self%>.txdays(index).value)*8.0
			<%=self%>.nowTX(index).value = cdbl(<%=self%>.tx_hr(index).value)-cdbl(<%=self%>.allJiaE(index).value)
		end if 
	else
		<%=self%>.txdays(index).value = <%=self%>.old_txdays(index).value 
		<%=self%>.tx_hr(index).value = cdbl(<%=self%>.txdays(index).value)*8.0
		<%=self%>.nowTX(index).value = cdbl(<%=self%>.tx_hr(index).value)-cdbl(<%=self%>.allJiaE(index).value)
	end if  
	if <%=self%>.nowTX(index).value > "0"  then
		<%=self%>.njAmt(index).value=formatnumber(cdbl(<%=self%>.nowTX(index).value)*3*cdbl(<%=self%>.hh_money(index).value),0)
	else	
		<%=self%>.njAmt(index).value = 0 
	end if
end function  

function insMemo(index)
tp=<%=self%>.totalpage.value
	cp=<%=self%>.CurrentPage.value
	rc=<%=self%>.RecordInDB.value
	YYMM = <%=self%>.YYMM.value
	open "<%=self%>.memo.asp?index="& index &"&currentpage=" & cp &"&yymm=" & yymm  , "_blank" , "top=10, left=10, width=450, height=450, scrollbars=yes"
end function 

function del(index) 
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del" 		
		open "<%=self%>.back.asp?func=del&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	else
		<%=self%>.op(index).value="no"  
		open "<%=self%>.back.asp?func=no&index="& index &"&CurrentPage="& <%=CurrentPage%> , "Back"
	end if 	 	
	'parent.best.cols="70%,30%"
end function 

function BACKMAIN()	
	open "../main.asp" , "_self"
end function   

function oktest(N)	
	tp=<%=self%>.totalpage.value 
	cp=<%=self%>.CurrentPage.value 
	rc=<%=self%>.RecordInDB.value 
	'open "empworkB.fore.asp?empautoid="& N &"&yymm="&"<%=calcmonth%>", "_self" 
	open "empworkB.fore.asp?empautoid="& N &"&YYMM="&"<%=calcmonth%>" &"&Ftotalpage=" & tp &"&Fcurrentpage=" & cp &"&FRecordInDB=" & rc , "_self" 
end function   

FUNCTION CLR()
	OPEN "<%=SELF%>.ASP" , "_self"
END FUNCTION 

function strchg(a)
	if a=1 then 
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value) 
		'IF TRIM(<%=self%>.empid.value)<>"" THEN 
			<%=SELF%>.totalpage.VALUE=0
			<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP?TOTALPAE=0"
			<%=SELF%>.SUBMIT()
		'END IF 
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function   

function go()
	<%=self%>.action="<%=self%>.upd.asp" 
	<%=self%>.submit()
end function 

function tp01clk()
	'sd.item(0).style.visibility="hidden"
	set objall=document.all.sd
   strtempno="文件中 ID 名稱是 sd 的總數：" & objall.length
   strtemp="文件中 ID 是 sd 的標籤："
   for inti=0 to objall.length-1
     'strtemp=strtemp & objall.item(inti).tagname & "  "
		 objall.item(inti).style.visibility="hidden"
   next
   strtemp=strtempno & chr(10) & strtemp & chr(10) 
   strtemp=strtemp & "第三個 test 的內容：" & objall.item(2).innerHTML
   alert(strtemp) 	
end function 	
</script>


<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->

<%
Set conn = GetSQLServerConnection()   
self="YEBE0501"  


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
planyear = request("planyear")

gTotalPage = 5
PageRec = 15    'number of records per page
TableRec = 25    'number of fields per record     

if trim(planyear)<>"" then 
	sql="select a.*, isnull(b.cnt,0) cnt from  "&_	
		"(select * from studyplan where yy='"& planyear &"' and isnull(status,'')<>'D'  ) a "&_
		"left join ( Select ssno, count(*) cnt from empstudy where isnull(Status,'')<>'D' group by ssno ) b on b.ssno=a.ssno "&_
		"order by a.ssno "
else
	sql="select '0' as cnt, * from studyplan where yy='1' order by ssno "
end if 	

'response.write sql 
'response.end 

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 	 	  		
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL, conn, 3, 3  
	IF NOT RS.EOF THEN
		'PageRec = rs.RecordCount 
		rs.PageSize = PageRec
		RecordInDB = rs.RecordCount
		TotalPage = rs.PageCount
		gTotalPage = TotalPage+3
	END IF 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then	
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = trim(rs("yy"))
				tmpRec(i, j, 2) = trim(rs("ssno"))
				tmpRec(i, j, 3) = trim(rs("studyName"))
				tmpRec(i, j, 4) = trim(rs("T1"))
				tmpRec(i, j, 5) = trim(rs("T2"))
				tmpRec(i, j, 6) = trim(rs("T3"))
				tmpRec(i, j, 7) = trim(rs("T4"))
				tmpRec(i, j, 8) = trim(rs("T5"))
				tmpRec(i, j, 9) = trim(rs("T6"))
				tmpRec(i, j, 10) = trim(rs("T7"))
				tmpRec(i, j, 11) = trim(rs("T8"))
				tmpRec(i, j, 12) = trim(rs("T9"))
				tmpRec(i, j, 13) = trim(rs("T10"))
				tmpRec(i, j, 14) = trim(rs("T11"))
				tmpRec(i, j, 15) = trim(rs("T12"))
				tmpRec(i, j, 16) = trim(rs("pcnt"))
				tmpRec(i, j, 17) = trim(rs("hhour"))
				tmpRec(i, j, 18) = trim(rs("amt"))
				tmpRec(i, j, 19) = trim(rs("dm"))
				tmpRec(i, j, 20) = trim(rs("nw"))
				tmpRec(i, j, 21) = trim(rs("memo"))				
				tmpRec(i, j, 22) = trim(rs("aid"))
				tmpRec(i, j, 23) = trim(rs("cnt"))				
				rs.movenext
			else
				exit for			
			end if 
	
			if rs.EOF then
				rs.Close
				Set rs = nothing
				exit for
			 end if
		next
	next 
	Session("YEBE0501") = tmpRec
else
	TotalPage = (request("TotalPage"))
	gTotalPage = (request("gTotalPage"))
	StoreToSession()
	CurrentPage = (request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBE0501")

	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) < cint(gTotalPage) then
			     CurrentPage = CurrentPage + 1
			  else
			  	 CurrentPage = TotalPage
		      end if
	     Case "END"
		      CurrentPage = gTotalPage
	     Case Else
		      CurrentPage = 1
	end Select 	
end if 	 
conn.close
set conn=nothing 
%>

<html> 
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<!-- #include file="../Include/func.inc" -->
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function f()
    if <%=self%>.planyear.value="" then 
    	<%=self%>.planyear.focus()   
    end if 	
    '<%=self%>.country.SELECT()
end function     

function planyearchg()
	if <%=self%>.planyear.value<>"" then 
		<%=self%>.totalpage.value="0"
		<%=self%>.submit()
	end if	
end function 
-->
</SCRIPT>   
</head> 
<body  onkeydown="enterto()"  onload=f() >
<form name="<%=self%>" method="post" action="<%=self%>.Fore.asp"  >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td>計畫年度 THỐNG KÊ NĂM</td>
									<td><input type="text" style="width:100px" name=planyear size=10 onblur="planyearchg()" value="<%=planyear%>"></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table id="myTableGrid" width="98%">     	
								<tr bgcolor=#e4e4e4>            
									<td align=center rowspan=2 width=30>刪除<br>Xóa</td> 
									<td align=center rowspan=2 width=30>STT</td>        	
									<td align=center rowspan=2 width=120>課程名稱<br>Tên môn học</td>
									<td align=center colspan=12 width=360>訓練月份<br>Tháng đào tạo</td>
									<td align=center rowspan=2 width=40>人數<br>Số người</td>
									<td align=center rowspan=2 width=40>時數<br>Số giờ</td>
									<td align=center rowspan=2 width=40>內外<br>Trong ngoài</td>
									<td align=center rowspan=2 width=60>費用<br>Chi phí</td>
									<td align=center rowspan=2 width=40>幣別<br>Loại tiền</td>
									<td align=center rowspan=2 width=120>備註<br>Ghi chú</td>
								</tr>
								<tr bgcolor=#e4e4e4>
									<td width=30 align=center>1</td>
									<td width=30  align=center>2</td>
									<td width=30 align=center>3</td>
									<td width=30  align=center>4</td>
									<td width=30  align=center>5</td>
									<td width=30  align=center>6</td>
									<td width=30  align=center>7</td>
									<td width=30  align=center>8</td>
									<td width=30  align=center>9</td>
									<td width=30  align=center>10</td>
									<td width=30  align=center>11</td>
									<td width=30  align=center>12</td>
								</tr>
								<%for CurrentRow = 1 to pagerec        
									 if CurrentRow mod 2 = 0 then
										wkcolor="#ffffff"
									 else
										wkcolor="LightGoldenrodYellow"
									 end if
								%>        	
									<tr bgcolor="<%=wkcolor%>">	        	
										<td  >
											<%if tmpRec(CurrentPage, CurrentRow, 23)>0 then %>
												<input type=hidden  name=func  >
												<input type=hidden name=op value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>" >
											<%else%>
												<%if tmpRec(CurrentPage, CurrentRow, 0)="del" then %>
													<input type=checkbox  name=func checked onclick="delchg(<%=CurrentRow-1%>)" >
													<input type=hidden name=op value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>" >
												<%else%>
													<input type=checkbox  name=func  onclick="delchg(<%=CurrentRow-1%>)"  >
													<input type=hidden name=op value="<%=tmpRec(CurrentPage, CurrentRow, 0)%>" >
												<%end if%>	
											<%end if%>	
										</td>
										<td  ><%=right("00"&(currentpage-1)*15+CurrentRow,2)%></td>        		
										<td >
											<input name=studyName size=28 class=inputbox   value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>" onchange="datachg(<%=currentRow-1%>)" >
											<input type=hidden name=ssno class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>"  >
										</td>
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 4)="Y" then %>
												<input type=checkbox name=m1  checked onclick="mchk1(<%=currentRow-1%>)" >
											<%else%>
												<input  type=checkbox name=m1  onclick="mchk1(<%=currentRow-1%>)" >
											<%end if%>	
										</td>        	
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 5)="Y" then %>
												<input  type=checkbox name=m2  checked  onclick=mchk2(<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m2  onclick=mchk2(<%=currentRow-1%>) >
											<%end if%>	
										</td>        	
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 6)="Y" then %>
												<input type=checkbox name=m3  checked  onclick=mchk3(<%=currentRow-1%>) >
											<%else%>
												<input  type=checkbox name=m3  onclick=mchk3(<%=currentRow-1%>) >
											<%end if%>	
										</td>        	
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 7)="Y" then %>
												<input type=checkbox name=m4  checked  onclick=mchk4(<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m4  onclick=mchk4(<%=currentRow-1%>) >
											<%end if%>	
										</td>        	
										<td  >
											<%if tmpRec(CurrentPage, CurrentRow, 8)="Y" then %>
												<input type=checkbox name=m5  checked  onclick=mchk5(<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m5  onclick=mchk5(<%=currentRow-1%>) >
											<%end if%>	
										</td>        					
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 9)="Y" then %>
												<input type=checkbox name=m6  checked  onclick=mchk6(<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m6  onclick=mchk6(<%=currentRow-1%>) >
											<%end if%>	
										</td>    
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 10)="Y" then %>
												<input type=checkbox name=m7  checked  onclick=mchk7(<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m7 onclick=mchk7(<%=currentRow-1%>) ) >
											<%end if%>	
										</td>            		
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 11)="Y" then %>
												<input type=checkbox name=m8  checked onclick=mchk(8<%=currentRow-1%>) >
											<%else%>
												<input type=checkbox name=m8 onclick=mchk8(<%=currentRow-1%>) >
											<%end if%>	
										</td>
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 12)="Y" then %>
												<input type=checkbox name=m9  checked onclick=mchk9(<%=currentRow-1%>)>
											<%else%>
												<input type=checkbox name=m9 onclick=mchk9(<%=currentRow-1%>)>
											<%end if%>	
										</td>
										<td   >
											<%if tmpRec(CurrentPage, CurrentRow, 13)="Y" then %>
												<input type=checkbox name=m10  checked onclick=mchk10(<%=currentRow-1%>)>
											<%else%>
												<input type=checkbox name=m10 onclick=mchk10(<%=currentRow-1%>) >
											<%end if%>	
										</td>				
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 14)="Y" then %>
												<input type=checkbox name=m11  checked onclick=mchk11(<%=currentRow-1%>)>
											<%else%>
												<input type=checkbox name=m11 onclick=mchk11(<%=currentRow-1%>)>
											<%end if%>	
										</td>
										<td>
											<%if tmpRec(CurrentPage, CurrentRow, 15)="Y" then %>
												<input type=checkbox name=m12  checked onclick=mchk12(<%=currentRow-1%>)>
											<%else%>
												<input type=checkbox name=m12 onclick=mchk12(<%=currentRow-1%>)>
											<%end if%>	
											<input type=hidden value="" name="T1">
											<input type=hidden value="" name="T2">
											<input type=hidden value="" name="T3">
											<input type=hidden value="" name="T4">
											<input type=hidden value="" name="T5">
											<input type=hidden value="" name="T6">
											<input type=hidden value="" name="T7">
											<input type=hidden value="" name="T8">
											<input type=hidden value="" name="T9">
											<input type=hidden value="" name="T10">
											<input type=hidden value="" name="T11">
											<input type=hidden value="" name="T12">
										</td>				
										<td ><input name=pcnt size=3 class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" style='text-align:right' onchange="datachg(<%=currentRow-1%>)" ></td>
										<td ><input name=hhour size=3 class=inputbox   value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>"  style='text-align:right' onchange="datachg(<%=currentRow-1%>)" ></td>	        	
										<td  >
											<select name=nw class=inputbox8 onchange="datachg(<%=currentRow-1%>)" >
												<option value=""></option>
												<option value="N" <%if tmpRec(CurrentPage, CurrentRow, 20)="N" then %>selected<%end if%> >內Nội</option>
												<option value="W" <%if tmpRec(CurrentPage, CurrentRow, 20)="W" then %>selected<%end if%> >外Ngoại</option>
											</seelct>
										</td>	
										<td ><input name=amt size=10 class=inputbox8   value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" style='text-align:right' onchange="datachg(<%=currentRow-1%>)" ></td>	        	        	
										<td  >
											<select name=dm class=inputbox8 onchange="datachg(<%=currentRow-1%>)" >
												<option value=""></option>
												<option value="VND"  <%if tmpRec(CurrentPage, CurrentRow, 19)="VND" then %>selected<%end if%> >VND</option>
												<option value="USD" <%if tmpRec(CurrentPage, CurrentRow, 19)="USD" then %>selected<%end if%>>USD</option>
											</seelct>
										</td>	
										<td  ><input name=memo size=25 class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 21)%>"   onchange="datachg(<%=currentRow-1%>)"  ></td>
									</tr>        	
								<%next%>
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt">
								<tr>
									<td align="CENTER" height=40 width=60%>
									PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%><BR>
									<% If CurrentPage > 1 Then %>
										<input type="submit" name="send" value="FIRST" class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="BACK" class="btn btn-sm btn-outline-secondary">
									<% Else %>
										<input type="submit" name="send" value="FIRST" disabled class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="BACK" disabled class="btn btn-sm btn-outline-secondary">
									<% End If %>
									<% If cint(CurrentPage) < cint(gTotalPage) Then %>
										<input type="submit" name="send" value="NEXT" class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="END" class="btn btn-sm btn-outline-secondary">
									<% Else %>
										<input type="submit" name="send" value="NEXT" disabled class="btn btn-sm btn-outline-secondary">
										<input type="submit" name="send" value="END" disabled class="btn btn-sm btn-outline-secondary">
									<% End If %>
									</td>
									<td><BR>
										<%if UCASE(session("mode"))="W" then%>
											<input type="button" name="send" value="CONFRIM" onclick="GO()" class="btn btn-sm btn-danger">
											<input type="reset" name="send" value="CANCEL" class="btn btn-sm btn-outline-secondary" >
										<%end if%>
									</td>
								</TR>
							</TABLE>
						</td>
					</tr>
				</table>
			
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("YEBE0501")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 3) = request("studyname")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 4) = request("T1")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 5) = request("T2")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("T3")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 7) = request("T4")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 8) = request("T5")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 9) = request("T6")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 10) = request("T7")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 11) = request("T8")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 12) = request("T9")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 13) = request("T10")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 14) = request("T11")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 15) = request("T12")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 16) = request("pcnt")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 17) = request("hhour")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 18) = request("amt")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 19) = request("dm")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 20) = request("nw")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 21) = request("memo")(CurrentRow)
	next 
	Session("YEBE0501") = tmpRec
End Sub
%>  

<script language=vbs> 

function delchg(index)
	if <%=self%>.func(index).checked=true then 
		<%=self%>.op(index).value="del" 
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> & "&func=del" , "Back"
	else
		<%=self%>.op(index).value="" 
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> & "&func=no" , "Back"
	end if
end function 
function dataclick(a)
    if a = 1 then       
        open "empbasic/empbasic.asp" , "_self"
    elseif a = 2 then       
        open "empfile/empfile.asp" , "_self"
    elseif a = 3 then       
        open "empworkHour/empwork.asp" , "_self"    
    elseif a = 4 then       
        open "holiday/empholiday.asp" , "_self" 
    elseif a = 5 then       
        open "AcceptCaTime/main.asp" , "_self"              
    elseif a = 6 then       
        open "../report/main.asp" , "_self"     
    end if      
end function  

function datachg(index)
	studyname_str = escape(ucase(trim(<%=self%>.studyName(index).value)))
	amt_str = <%=self%>.amt(index).value
	dm_str = <%=self%>.Dm(index).value
	pcnt_str = <%=self%>.pcnt(index).value
	hhour_str = <%=self%>.hhour(index).value
	nw_str = <%=self%>.nw(index).value
	memo_str = escape(ucase(trim(<%=self%>.memo(index).value)))
	
	open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &_
		 "&code1=" & studyname_str &"&code2=" & pcnt_str &"&code3=" & hhour_str &_
		 "&code4=" & nw_str &"&code5=" & amt_str &"&code6="& dm_str  &_
		 "&code7=" & memo_str&  "&func=datachg" , "Back" 

'parent.best.cols="70%,30%"		 
end  function 

function strchg(a)
    if a=1 then 
        <%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
    elseif a=2 then     
        <%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
    end if  
end function 
    
function go() 
    <%=self%>.action="<%=self%>.upd.asp"
    <%=self%>.submit() 
end function   

function mchk1(index)
	if <%=self%>.m1(index).checked=true then
		<%=self%>.t1(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&CurrentPage=" & <%=CurrentPage%> &"&func=T1Y" , "Back"
	else
		<%=self%>.t1(index).value=""
		open "<%=self%>.back.asp?index="&index &"&CurrentPage=" & <%=CurrentPage%> &"&func=T1N" , "Back"
	end if 	
	'parent.best.cols="70%,30%"
end function  

function mchk2(index)
	if <%=self%>.m2(index).checked=true then
		<%=self%>.t2(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T2Y" , "Back"
	else
		<%=self%>.t2(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T2N" , "Back"
	end if 	
end function   

function mchk3(index)
	if <%=self%>.m3(index).checked=true then
		<%=self%>.t3(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T3Y" , "Back"
	else
		<%=self%>.t3(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T3N" , "Back"
	end if 	
end function  
function mchk4(index)
	if <%=self%>.m4(index).checked=true then
		<%=self%>.t4(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T4Y" , "Back"
	else
		<%=self%>.t4(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T4N" , "Back"
	end if 	
end function  
function mchk5(index)
	if <%=self%>.m5(index).checked=true then
		<%=self%>.t5(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T5Y" , "Back"
	else
		<%=self%>.t5(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T5N" , "Back"
	end if 	
end function  
function mchk6(index)
	if <%=self%>.m6(index).checked=true then
		<%=self%>.t6(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T6Y" , "Back"
	else
		<%=self%>.t6(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T6N" , "Back"
	end if 	
end function  
function mchk7(index)
	if <%=self%>.m7(index).checked=true then
		<%=self%>.t7(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T7Y" , "Back"
	else
		<%=self%>.t7(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T7N" , "Back"
	end if 	
end function  
function mchk8(index)
	if <%=self%>.m8(index).checked=true then
		<%=self%>.t8(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T8Y" , "Back"
	else
		<%=self%>.t8(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T8N" , "Back"
	end if 	
end function  
function mchk9(index)
	if <%=self%>.m9(index).checked=true then
		<%=self%>.t9(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T9Y" , "Back"
	else
		<%=self%>.t9(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T9N" , "Back"
	end if 	
end function  
function mchk10(index)
	if <%=self%>.m10(index).checked=true then
		<%=self%>.t10(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T10Y" , "Back"
	else
		<%=self%>.t10(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T10N" , "Back"
	end if 	
end function  
function mchk11(index)
	if <%=self%>.m11(index).checked=true then
		<%=self%>.t11(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T11Y" , "Back"
	else
		<%=self%>.t11(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T11N" , "Back"
	end if 	
end function  
function mchk12(index)
	if <%=self%>.m12(index).checked=true then
		<%=self%>.t12(index).value="Y"
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T12Y" , "Back"
	else
		<%=self%>.t12(index).value=""
		open "<%=self%>.back.asp?index="&index &"&currentpage=" & <%=CurrentPage%> &"&func=T12N" , "Back"
	end if 	
end function  
    

'*******檢查日期*********************************************
FUNCTION date_change(a) 

if a=1 then
    INcardat = Trim(<%=self%>.indat1.value)         
elseif a=2 then
    INcardat = Trim(<%=self%>.indat2.value)
end if      
            
IF INcardat<>"" THEN
    ANS=validDate(INcardat)
    IF ANS <> "" THEN
        if a=1 then
            Document.<%=self%>.indat1.value=ANS         
        elseif a=2 then
            Document.<%=self%>.indat2.value=ANS                 
        end if      
    ELSE
        ALERT "EZ0067:輸入日期不合法 !!" 
        if a=1 then
            Document.<%=self%>.indat1.value=""
            Document.<%=self%>.indat1.focus()
        elseif a=2 then
            Document.<%=self%>.indat2.value=""
            Document.<%=self%>.indat2.focus()
        end if      
        EXIT FUNCTION
    END IF
         
ELSE
    'alert "EZ0015:日期該欄位必須輸入資料 !!"      
    EXIT FUNCTION
END IF 
   
END FUNCTION 

function gotstudyplan()
    ncols="studyGroup" 
    open "getstudyPlan.asp?pself="& "<%=self%>" &"&ncols="& ncols , "Back" 
    parent.best.cols="50%,50%" 
    
    'open "Getempdata.asp?pself="& "<%=self%>" &"&index=" & index &"&ncols="& ncols , "Back"   
end function 
</script> 
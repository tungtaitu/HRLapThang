<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<!-- #include file="../ADOINC.inc" -->
<!-- #include file="../Include/func.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
'on error resume next   
session.codepage="65001"
SELF = "EMPHOLIDAYB"
 
Set conn = GetSQLServerConnection()	  
Set rs = Server.CreateObject("ADODB.Recordset")   
DAT1 = REQUEST("DAT1")
DAT2 = REQUEST("DAT2")
whsno = trim(request("whsno"))
groupid = trim(request("groupid"))
country = trim(request("country"))  
QUERYX = trim(request("empid1"))  

unitno = trim(request("unitno"))
zuno = trim(request("zuno"))
job = trim(request("job")) 
jb = trim(request("jb")) 
ym1= trim(request("ym1")) 
ym2= trim(request("ym2")) 

gTotalPage = 1
PageRec = 16    'number of records per page
TableRec = 30    'number of fields per record  

if dat1="" and dat2="" and whsno="" and groupid="" and country="" and QUERYX="" then 
	sql="select * from empfile where empid='XX' "
else
	SQL="SELECT  A.JIATYPE,  CONVERT(CHAR(10), A.DATEUP, 111) DATEUP , A.TIMEUP, convert(char(10) , A.DATEDOWN , 111) datedown, "
	SQL=SQL&"A.TIMEDOWN , A.HHOUR, A.MEMO AS JIAMEMO  , a.autoid as jiaid, isnull(a.xjsts,'')xjsts, B.*  , isnull(c.sys_value,'') as jia_str  FROM   "
	SQL=SQL&"( SELECT * FROM EMPHOLIDAY   ) A  "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM view_empfile ) B ON B.EMPID = A.EMPID  	 "
	SQL=SQL&"LEFT JOIN ( SELECT * FROM basicCode where func='JB'  ) c  on c.sys_type = a.JIATYPE  "
	SQL=SQL&"WHERE 1=1  " 	
	SQL=SQL&"and country like  '"& country &"%'  "
	SQL=SQL&"AND whsno like '"& whsno &"%' and unitno like '"& unitno &"%'  and groupid like '"& groupid &"%'  " 
	SQL=SQL&"and zuno like '"& zuno &"%' and a.jiatype like '"& jb &"%' and b.empid like '%"& QUERYX &"%'  "
	
	IF DAT1<>"" and DAT2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(10), A.DATEUP, 111) BETWEEN '"& DAT1 &"' AND '"& DAT2 &"' " 
	END IF 
	IF ym1<>"" and ym2<>"" then  
	 	sql=sql& "and CONVERT(CHAR(6), A.DATEUP, 112) BETWEEN '"& ym1 &"' AND '"& ym2 &"' " 
	END IF 
	SQL=SQL&"order by b.empid, A.DATEUP , a.jiaType "  
end if 	
'response.write sql 
'RESPONSE.END  
if request("TotalPage") = "" or request("TotalPage") = "0" then 
	CurrentPage = 1	
	rs.Open SQL, conn, 3, 1
	IF NOT RS.EOF THEN 
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
			tmpRec(i, j, 1) = trim(rs("empid"))
			tmpRec(i, j, 2) = trim(rs("empnam_cn"))
			tmpRec(i, j, 3) = trim(rs("empnam_vn"))
			tmpRec(i, j, 4) = rs("country")
			tmpRec(i, j, 5) = rs("nindat")
			tmpRec(i, j, 6) = rs("job")				
			tmpRec(i, j, 7) = rs("whsno")	 
			tmpRec(i, j, 8) = rs("unitno")	 
			tmpRec(i, j, 9)	=RS("groupid") 
			tmpRec(i, j, 10)=RS("zuno") 				
			tmpRec(i, j, 11)=RS("whsno") 	
			tmpRec(i, j, 12)="" 'RS("ustr") 	
			tmpRec(i, j, 13)=RS("gstr") 	
			tmpRec(i, j, 14)=RS("zstr") 	
			tmpRec(i, j, 15)=RS("jstr") 	
			tmpRec(i, j, 16)="" 'RS("cstr")
			tmpRec(i, j, 17)=RS("DATEUP")
			tmpRec(i, j, 18)=RS("TIMEUP")
			tmpRec(i, j, 19)=RS("DATEDOWN")
			tmpRec(i, j, 20)=RS("TIMEDOWN")
			tmpRec(i, j, 21)=RS("JIAMEMO")
			tmpRec(i, j, 22)=RS("JIATYPE") 
			tmpRec(i, j, 23)=RS("hhour") 
			tmpRec(i, j, 24)=RS("jiaID")  
			tmpRec(i, j, 25)=RS("jia_str")  
			tmpRec(i, j, 26)=RS("DATEUP") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEUP"))) , 1 )  
			tmpRec(i, j, 27)=RS("DATEDOWN") &" "&mid("日一二三四五六",weekday(cdate(rs("DATEDOWN"))) , 1 )  
			tmpRec(i, j, 28)=RS("xjsts") 
			 
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
	Session("EMPHOLIDAYB") = tmpRec	
else    
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("EMPHOLIDAYB")
	
	Select case request("send") 
	     Case "FIRST"
		      CurrentPage = 1			
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then 
			     CurrentPage = CurrentPage - 1				
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(TotalPage) then 
			     CurrentPage = CurrentPage + 1 
		      end if			
	     Case "END"
		      CurrentPage = TotalPage 			
	     Case Else 
		      CurrentPage = 1	
	end Select 
end if   


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

function datachg()
	<%=SELF%>.totalpage.VALUE=0
	<%=self%>.action="<%=SELF%>.fore.asp"
	<%=self%>.submit
end function 

-->
</SCRIPT>  
</head>   
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=SELF%>.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>"> 	 

	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

				<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td >							
							<table class="txt" cellpadding=3 cellspacing=3> 								  
								<tr>
									<td align="right">年月<br><font class="txt8">Thang Nam</font></td>
									<Td nowrap>
										<INPUT type="text" style="width:100px" NAME=ym1 VALUE="<%=ym1%>" maxlength=6  >~
										<INPUT type="text" style="width:100px" NAME=ym2 VALUE="<%=ym2%>" maxlength=6 onblur="gos()">
									</td>		
									<td align="right">日期<br><font class="txt8">Ngay</font></td>
									<td nowrap>
										<INPUT type="text" style="width:120px" NAME=DAT1 VALUE="<%=DAT1%>" maxlength=10   onblur="date_change(1)" >~
										<INPUT type="text" style="width:120px" NAME=DAT2 VALUE="<%=DAT2%>"  maxlength=10  onblur="date_change(2)">
									</td>
								</tr>
								<TR >
									<TD nowrap align=right>假別<br><font class="txt8">Loai phep</font></TD>
									<TD > 
										<select name=JB   onchange="datachg()" style="width:120px">
											<option value="">---</option>
											<%SQL="SELECT * FROM BASICCODE WHERE FUNC='jb' ORDER BY SYS_TYPE "
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF  
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if jb=rst("sys_type") then%>selected<%end if%> ><%=RST("SYS_TYPE")%>-<%=RST("SYS_VALUE")%></option>				 
											<%
											RST.MOVENEXT
											WEND 
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>  
									<TD nowrap align=right >員工編號<br><font class="txt8">So the</font></TD>
									<TD >
										<input type="text" style="width:100px" name=empid1 maxlength=6 ONBLUR=strchg(1) VALUE="<%=QUERYX%>"> 
									</TD> 	 		
								</TR>
								<tr>
									<td colspan=4>
										<input type="checkbox" name="sall"  onclick="selectAll()" >全選 (全部刪除)
										<input type="hidden" name="delAll" value="N">
									</td>
								</tr>
							</TABLE>
						</td>
					</tr>					
					<tr>
						<td>
							<table id="myTableGrid" width="98%">
								<tr BGCOLOR="LightGrey" height="35px" >
									<TD width=30 nowrap align=center >刪除<br><font class="txt8">Xoa</font></TD> 
									<TD width=50 nowrap align=center >工號<br><font class="txt8">So The</font></TD> 		
									<TD width=190 nowrap align=center >姓名<br><font class="txt8">Ho ten</font></TD>
									<TD align=center  >假別<br><font class="txt8">loai phep</font></TD>
									<TD width=80 align=center nowrap >日期(起)<br><font class="txt8">Ngay(tu)</font></TD>
									<TD align=center  >時間(起)<br><font class="txt8">Thoi gian(tu)</font></TD>
									<TD width=80 align=center nowrap >日期(迄)<br><font class="txt8">Ngay(Den)</font></TD>
									<td align=center  >時間(迄)<br><font class="txt8">Thoi gian(den)</font></td>
									<td align=center  >時數<br><font class="txt8">So gio</font></td>
									<td align=center >事由<br><font class="txt8">Ly do</font></td>		
								</tr>
								 
								<%for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN 
										WKCOLOR="LavenderBlush"
									ELSE
										WKCOLOR=""
									END IF 	 
									'if tmpRec(CurrentPage, CurrentRow, 1) <> "" then 
								%>
								<TR BGCOLOR=<%=WKCOLOR%> > 	
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>		
											<%IF tmpRec(CurrentPage, CurrentRow, 0)="del" THEN  %>
												<INPUT type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)" checked >
											<%ELSE%>	
												<INPUT type=checkbox name=func value=del onclick="del(<%=CurrentRow - 1%>)"   >
											<%END IF%>	
											<INPUT TYPE=HIDDEN NAME=OP >
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=FUNC  >
											<INPUT TYPE=HIDDEN NAME=OP   >
										<%END IF %>
									</TD>
									<TD align=center>
										<%=tmpRec(CurrentPage, CurrentRow, 1)%>
										<INPUT TYPE=HIDDEN NAME=EMPID VALUE="<%=tmpRec(CurrentPage, CurrentRow, 1)%>">
									</TD> 		
									<TD>
										<%=tmpRec(CurrentPage, CurrentRow, 2)%>&nbsp;
										<font class=txt8><%=tmpRec(CurrentPage, CurrentRow, 3)%></font>
									</TD>
									<TD>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN 
										%>
											<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>" >
											<INPUT type="text" style="width:100%" NAME=HOLIDASTR value="<%=tmpRec(CurrentPage, CurrentRow, 22)%>&nbsp;<%=tmpRec(CurrentPage, CurrentRow, 25)%>" class=readonly  readonly> 	 			 
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=HOLIDAY_TYPE  >	
											<INPUT TYPE=HIDDEN NAME=HOLIDASTR >			
										<%END IF %>
									</TD>
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
											<input type="text" style="width:100%" name=HHDAT1 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 26)%>" > 				
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=HHDAT1  >								
										<%END IF %>
									</TD>
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
											<input type="text" style="width:100%" name=HHTIM1 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 18)%>" style="text-align:center" >
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=HHTIM1  >				
										<%END IF %>
									</TD>
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	 				
											<input type="text" style="width:100%" name=HHDAT2 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 27)%>" >
										<%ELSE%>					
											<INPUT TYPE=HIDDEN NAME=HHDAT2  >			
										<%END IF %>
									</TD> 
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
											<input type="text" style="width:100%" name=HHTIM2 class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 20)%>" style="text-align:center">
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=HHTIM2  >				
										<%END IF %>
									</TD>
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN %>	
											<input type="text" style="width:100%" name=toth class=readonly readonly   value="<%=tmpRec(CurrentPage, CurrentRow, 23)%>" style="text-align:right">
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=toth  >				
										<%END IF %>
									</TD> 
									<TD align=center>
										<%IF tmpRec(CurrentPage, CurrentRow, 1)<>"" THEN
											reson = tmpRec(CurrentPage, CurrentRow, 21) 
											if tmpRec(CurrentPage, CurrentRow, 28)="C" then reson="不扣全勤,"&reson
										%>	
											<input type="text" style="width:100%" name=JIAMEMO class=readonly readonly   value="<%=reson%>" >
										<%ELSE%>	
											<INPUT TYPE=HIDDEN NAME=JIAMEMO  >				
										<%END IF %>
									</TD> 		
								</TR>
								<%next%> 	 
								
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="txt" cellpadding=3 cellspacing=3>
								<tr>
									<td align="CENTER" height=40 width=70%>									
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
										<input type="BUTTON" name="send" value="(Y)Confirm" class="btn btn-sm btn-danger" onclick="GO()" >
										<input type="BUTTON" name="send" value="(N)Cancel" class="btn btn-sm btn-outline-secondary" ONCLICK="CLR()">		
									</td>		
								</TR>
							</TABLE>  
							<input type=hidden name=func >
							<input type=hidden name=op >
							<input type=hidden name=empid >
						</td>
					</tr>
				</table>
			
</form>

</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("EMPHOLIDAYB")
	for CurrentRow = 1 to PageRec
		tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)	
	next 
	Session("EMPHOLIDAYB") = tmpRec
	
End Sub
%>  
<script language=vbscript> 
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
			<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP"
			<%=SELF%>.SUBMIT()
		'END IF 
	elseif a=2 then 	
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if 	
end function   

function go()
	<%=self%>.action="<%=self%>.updateDB.asp" 
	<%=self%>.submit()
end function  

function selectAll()
	if <%=self%>.sall.checked=true then 
		<%=self%>.delAll.value="Y"
	else
		<%=self%>.delAll.value="N"
	end if 
end function  

function gos()	
	<%=SELF%>.totalpage.VALUE=0
	<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP"
	<%=SELF%>.SUBMIT()
end function 

'*******檢查日期*********************************************
FUNCTION date_change(a)	

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)  		
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
end if		
		    
IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS			
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS		 			 
			if <%=self%>.dat1.value<>"" then 
				<%=SELF%>.totalpage.VALUE=0
				<%=SELF%>.ACTION="<%=SELF%>.FORE.ASP"
				<%=SELF%>.SUBMIT()
			end if 
		end if		
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!" 
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		end if		
		EXIT FUNCTION
	END IF
		 
ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!" 		
	EXIT FUNCTION
END IF 
   
END FUNCTION
	
</script>


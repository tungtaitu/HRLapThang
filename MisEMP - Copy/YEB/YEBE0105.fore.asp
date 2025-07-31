<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%
if session("netuser")="" then 
	response.write "使用者帳號為空!!請重新登入!!"
	response.end 
end if 	

SELF = "YEBE0105"

Set conn = GetSQLServerConnection()


nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
end if
 

FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END FUNCTION

whsno = request("whsno")
wbloai = request("wbloai")
cmnd = request("cmnd")
EMPID =request("EMPID")
soxe=request("soxe")
fac=request("fac")

gTotalPage = 1
PageRec = 20    'number of records per page
TableRec = 30    'number of fields per record

if whsno="" and wbloai="" and cmnd="" and EMPID="" and soxe="" and fac="" then 
	sql="select * from wbempfile where wbid='X' "
else
	sql="select isnull(b.incustsname,'') incustsname, convert(varchar(10),indat,111) as nindat, convert(varchar(10),outdat,111) as outdate , a.* "&_
		"from "&_
		"(Select *from wbempfile where isnull(status,'')<>'D' and isnull(wbid,'') like '%"&empid&"%' and wbwhsno like '"&whsno&"%' and loai  like'"&wbloai&"%'  "&_
		"and isnull(personid,'') like '%"&cmnd&"%' and isnull(soxe,'') like '%"&soxe&"%' and isnull(fac,'') like '%"&fac &"%' ) a  "&_
		"left join (select incustid, incustsname from [yfymis].dbo.ydbscust ) b on b.incustid = a.fac "&_
		"order by wbid " 
end if 
'response.write sql 
'response.end 
Set rs = Server.CreateObject("ADODB.Recordset")

if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1
	rs.Open sql, conn, 3, 3
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
				tmpRec(i, j, 1) = trim(rs("wbid"))
				tmpRec(i, j, 2) = trim(rs("wbname_cn"))
				tmpRec(i, j, 3) = trim(rs("wbname_vn"))
				tmpRec(i, j, 4) = rs("wbwhsno")
				tmpRec(i, j, 5) = rs("loai")
				tmpRec(i, j, 6) = rs("yy")
				tmpRec(i, j, 7) = rs("mm")
				tmpRec(i, j, 8) = rs("dd")
				tmpRec(i, j, 9)	=RS("age")
				tmpRec(i, j, 10)=RS("nindat")
				tmpRec(i, j, 11)=RS("lorry")
				tmpRec(i, j, 12)=RS("soxe")
				tmpRec(i, j, 13)=RS("job")
				tmpRec(i, j, 14)=RS("personid")
				tmpRec(i, j, 15)=RS("phone")
				tmpRec(i, j, 16)=RS("mobile")
				tmpRec(i, j, 17)=RS("addr")
				tmpRec(i, j, 18)=RS("fac")
				tmpRec(i, j, 19)=RS("wbmemo")				 
				tmpRec(i, j, 20)=RS("sex")
				tmpRec(i, j, 21)=RS("outdate")
				tmpRec(i, j, 22)=RS("outmemo") 			
				if rs("dd")<>"" then 
					borndat=cstr(rs("dd"))&"/"
				else
					borndat=""
				end if 	
				if rs("mm")<>"" then 
					borndat=borndat&cstr(rs("mm"))&"/"
				else
					borndat=""					
				end if 
				if rs("yy")<>"" then 
					borndat=borndat&cstr(rs("yy")) 
				else
					borndat=""					
				end if 		
				tmpRec(i, j, 23)=borndat
				
				if trim(rs("incustsname"))<>"" then 
					tmpRec(i, j, 24)=RS("incustsname")
				else
					tmpRec(i, j, 24)=RS("fac")
				end if 	 
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
	Session("YEBE0104B") = tmpRec
else
	TotalPage = cint(request("TotalPage"))
	'StoreToSession()
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	tmpRec = Session("YEBE0104B")

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


FUNCTION FDT(D)
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)

END FUNCTION 

 

%>

<html>

<head>

<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css"> 
</head>
<body  onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.fore.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE="<%=SESSION("NETUSER")%>">
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table cellspacing="1" cellpadding="1" class="table table-bordered table-sm bg-white text-secondary" bgcolor=black>
								<tr bgcolor=#ffffff height=35>
									<Td align=center bgcolor="#ffffff" width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand" ><a href="yebe0103.fore.asp">外包資料新增<br>tu lieu moi thau ngoai</a></td>
									<Td align=center bgcolor="#ffffff"  width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0104.fore.asp">外包資料維護<br>xoa/sua tu lieu thau ngoai</a></td>
									<Td align=center bgcolor="#66ccff"  width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#66ccff'"   style="cursor:hand"><a href="yebe0105.fore.asp">外包資料查詢<br>K.Tra tu lieu</a></td>
									<Td align=center bgcolor="#ffffff"  width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0103C.asp">照片上傳與新增<br>Update Photos</a></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table-borderless table-sm bg-white text-secondary">
								<TR height=35 >
									<TD  width=60 align=right>廠別<BR><font class=txt8>Xuong</font></TD>
									<TD   valign=top width=80>
										<select name=WHSNO  class="form-control form-control-sm mb-2 mt-2" onchange='gos()' style='width:70'>
											<option value="">請選擇廠別</option>
											<%
											if session("rights")="0" then 
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO'  ORDER BY SYS_TYPE "
											else
												SQL="SELECT * FROM BASICCODE WHERE FUNC='WHSNO' and sys_type='"& session("netwhsno") &"' ORDER BY SYS_TYPE "
											end if 	
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>" <%if whsno=rst("SYS_TYPE") then%>selected<%end if%> ><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											%>
										</SELECT>
										<%SET RST=NOTHING %>
									</TD>	
									<TD   align=right width=70>類別<BR><font class=txt8>Loai</font></TD>
									<TD  valign=top width=120 >
										<select name=wbloai  class="form-control form-control-sm mb-2 mt-2" onchange='gos()' style='width:90'  >
											<option value="">請選擇類別</option>
											<%				
											SQL="SELECT * FROM BASICCODE WHERE FUNC='WB'   ORDER BY SYS_TYPE "				
											SET RST = CONN.EXECUTE(SQL)
											WHILE NOT RST.EOF
											%>
											<option value="<%=RST("SYS_TYPE")%>"  <%if wbloai=rst("SYS_TYPE") then%>selected<%end if%>><%=RST("SYS_TYPE")%><%=RST("SYS_VALUE")%></option>
											<%
											RST.MOVENEXT
											WEND
											rst.close
											%>
										</SELECT>
										<%SET RST=NOTHING 
										conn.close
										set conn=nothing			
										%>
									</TD>				
									<TD  align=right height=25 width=95>身分證號<BR><font class=txt8>CMND</font></TD>
									<TD  valign=top><INPUT NAME="cmnd" SIZE=15 CLASS="form-control form-control-sm mb-2 mt-2"  value="<%=cmnd%>" onblur="dchg(1)"></TD>
								</tr>
								<tr>
									<TD  align=right height=25>編號<BR><font class=txt8>So The</font></TD>
									<TD  valign=top><INPUT NAME="EMPID" SIZE=10 CLASS="form-control form-control-sm mb-2 mt-2"    maxlength=5 value="<%=eid%>" onblur="dchg(2)"></TD>
									<TD    align=right height=25>車號<BR><font class=txt8>So Xe</font></TD>
									<TD  valign=top><INPUT NAME="soxe" SIZE=10 CLASS="form-control form-control-sm mb-2 mt-2"   value="<%=soxe%>" onblur="dchg(3)"></TD>
									<TD align=right height=25>供應商<BR><font class=txt8>NHÀ CUNG ỨNG</font></TD>
									<TD  valign=top><INPUT NAME="fac" SIZE=12 CLASS="form-control form-control-sm mb-2 mt-2"   value="<%=fac%>" onblur="dchg(4)"></TD>
									
								</tr> 
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<table class="table table-bordered table-sm bg-white text-secondary">
								<tr bgcolor=#e4e4e4 height=25>
									<td align=center width=30>STT</td>
									<td align=center width=50>編號<br><font class=txt8>so the</font></td>		
									<td align=center width=90>姓名<br><font class=txt8>Ho ten</font></td>
									<td align=center width=70 >身分證號<br><font class=txt8>CMND</font></td>
									<td align=center width=70>上班日<br><font class=txt8>NVX</font></td>
									<td align=center width=70>生日<br><font class=txt8>Ngay Sinh</font></td>
									<td align=center width=80>供應商<br><font class=txt8>Nha Cung Ung</font></td>
									<td align=center width=60>車號<br><font class=txt8>So xe</font></td>
									<td align=center width=50>職務<br><font class=txt8>chuc vu</font></td>
									<td align=center width=60>手機<br><font class=txt8>DTDD</font></td>
									<td align=center width=100>備註說明<br><font class=txt8>Ghi Chu</font></td>
								</tr>
								<%for CurrentRow = 1 to PageRec
									IF CurrentRow MOD 2 = 0 THEN
										WKCOLOR="#ffffff"  '"LavenderBlush"
									ELSE
										WKCOLOR="#ffffff"
									END IF
									if tmpRec(CurrentPage, CurrentRow, 1) <> "" then
								%>
									<TR BGCOLOR='<%=WKCOLOR%>' height=22 class=txt8 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"  onclick="showdata(<%=currentrow-1%>)"  >
										<Td align=center><%=(currentpage-1)*PageRec+currentrow%></td>
										<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 1)%></td>
										<Td ><%=tmpRec(CurrentPage, CurrentRow, 3)%><br><%=tmpRec(CurrentPage, CurrentRow, 2)%></td>
										<Td align=center><%=tmpRec(CurrentPage, CurrentRow, 14)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 10)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 23)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 24)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 12)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 13)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 16)%></td>
										<Td><%=tmpRec(CurrentPage, CurrentRow, 19)%></td>
										<input name=wbid type=hidden value="<%=trim(tmpRec(CurrentPage, CurrentRow, 1))%>">
									</tr>
									<%end if%>
								<%next%>
								<input name=wbid type=hidden >	
							</table>
						</td>
					</tr>
					<tr>
						<td align="center">
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr ALIGN=center>
								<td align="left" height=40  >	    
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
								<% End If %>&nbsp;
								PAGE:<%=CURRENTPAGE%> / <%=TOTALPAGE%> , COUNT:<%=RECORDINDB%> 
								</td>		
								<td>
									<input type=button name=btn value="(N)重新查詢K.Tra"  class="btn btn-sm btn-outline-secondary" onclick="gon()">
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
 
'-----------------enter to next field 
function getlorry()
	open "getlorry.asp", "Back"
	parent.best.cols="50%,50%"
end function 
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function showdata(index)
	wbidstr = <%=self%>.wbid(index).value
	open "<%=self%>.foregnd.asp?wbid="&wbidstr , "_self" 
end function 

function f()
	<%=self%>.whsno.focus()
	'<%=self%>.EMPID.select()
end function   

function dchg(a)
	select case a 
		case 1 
			if trim(<%=self%>.cmnd.value)<>"" then 
				gos()
			end if 	
		case 2 
			if trim(<%=self%>.empid.value)<>"" then 
				gos()
			end if 
		case 3 
			if trim(<%=self%>.soxe.value)<>"" then 
				gos()
			end if 
		case 4 
			if trim(<%=self%>.fac.value)<>"" then 
				gos()
			end if 
	end select											
				
end function 

function gos() 
	<%=self%>.totalpage.value="0"
	<%=self%>.action="<%=self%>.Fore.asp"
	<%=self%>.target="_self"
	<%=self%>.submit()
end function	

function goN()
	open "<%=self%>.fore.asp" , "_self"	
end function	

FUNCTION GO()
	if <%=self%>.whsno.value="" then 
		ALERT "請選擇廠別(Ko. co nhap vao Xuong)!!"
		<%=SELF%>.whsno.FOCUS()
		EXIT FUNCTION 
	end if 
	if <%=self%>.wbloai.value="" then 
		ALERT "請選擇類別(Ko. co nhap vao Loai)!!"
		<%=SELF%>.wbloai.FOCUS()
		EXIT FUNCTION 
	end if 	
	IF  <%=SELF%>.EMPID.VALUE="" THEN
		ALERT "請輸入編號(Ko. co nhap vao So )!!"
		<%=SELF%>.EMPID.FOCUS()
		EXIT FUNCTION 
	END IF 	 
	if <%=self%>.nam_vn.value="" then 
		ALERT "請輸入姓名(越)(Ko. co nhap vao Ho Ten(Viet)!!"
		<%=SELF%>.nam_vn.FOCUS()
		EXIT FUNCTION 
	end if 
		
	IF  <%=SELF%>.personID.VALUE="" THEN
		ALERT "請輸入身分證號(Ko. co nhap vao CMND )!!"
		<%=SELF%>.personID.FOCUS()
		EXIT FUNCTION 
	END IF 	 
	IF  <%=SELF%>.fac.VALUE="" THEN
		ALERT "請輸入供應商/車行(Ko. co nhap vao NHÀ CUNG ỨNG )!!"
		<%=SELF%>.fac.FOCUS()			
		EXIT FUNCTION 
	end if 		
	if <%=self%>.wbloai.value="01" then 
		IF  <%=SELF%>.fac.VALUE="" THEN
			ALERT "請輸入供應商/車行(Ko. co nhap vao NHÀ CUNG ỨNG )!!"
			<%=SELF%>.fac.FOCUS()			
			EXIT FUNCTION 
		end if 	
		IF  <%=SELF%>.soxe.VALUE="" THEN
			ALERT "請輸入車號(Ko. co nhap vao So Xe )!!"
			<%=SELF%>.soxe.FOCUS()			
			EXIT FUNCTION 
		end if 
	END IF	
 
	
	<%=SELF%>.ACTION="<%=self%>.upd.asp?act=EMPADDNEW"
	<%=SELF%>.SUBMIT
END FUNCTION

'*******檢查日期*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.indat.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.BHDAT.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.pduedate.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.vduedate.value)
elseif a=5 then
	INcardat = Trim(<%=self%>.outdat.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.indat.value=ANS
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=ANS
		elseif a=3 then
			Document.<%=self%>.pduedate.value=ANS
		elseif a=4 then
			Document.<%=self%>.vduedate.value=ANS
		elseif a=5 then
			Document.<%=self%>.outdat.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.indat.value=""
			Document.<%=self%>.indat.focus()
		elseif a=2 then
			Document.<%=self%>.BHDAT.value=""
			Document.<%=self%>.BHDAT.focus()
		elseif a=3 then
			Document.<%=self%>.pduedate.value=""
			Document.<%=self%>.pduedate.focus()
		elseif a=4 then
			Document.<%=self%>.vduedate.value=""
			Document.<%=self%>.vduedate.focus()
		elseif a=5 then
			Document.<%=self%>.outdat.value=""
			Document.<%=self%>.outdat.focus()
		end if
		EXIT FUNCTION
	END IF

ELSE
	'alert "EZ0015:日期該欄位必須輸入資料 !!"
	EXIT FUNCTION
END IF

END FUNCTION

'_________________DATE CHECK___________________________________________________________________

function validDate(d)
	if len(d) = 8 and isnumeric(d) then
		d = left(d,4) & "/" & mid(d, 5, 2) & "/" & right(d,2)
		if isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
	elseif len(d) >= 8 and isdate(d) then
			validDate = formatDate(d)
		else
			validDate = ""
		end if
end function

function formatDate(d)
		formatDate = Year(d) & "/" & _
		Right("00" & Month(d), 2) & "/" & _
		Right("00" & Day(d), 2)
end function
'________________________________________________________________________________________

FUNCTION CHKVALUE(N)
IF N=1 THEN
	IF TRIM(<%=SELF%>.BYY.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BYY.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BYY.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BYY.VALUE=""
			<%=SELF%>.BYY.FOCUS()
			EXIT FUNCTION
		ELSE
			<%=SELF%>.AGES.VALUE=CDBL(YEAR(DATE()))-CDBL(<%=SELF%>.BYY.VALUE) + 1
		END IF
	END IF
ELSEIF N=2 THEN
	IF TRIM(<%=SELF%>.BMM.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BMM.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BMM.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BMM.VALUE=""
			<%=SELF%>.BMM.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=3 THEN
	IF TRIM(<%=SELF%>.BDD.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.BDD.VALUE)=FALSE OR INSTR(1,<%=SELF%>.BDD.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.BDD.VALUE=""
			<%=SELF%>.BDD.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=4 THEN
	IF TRIM(<%=SELF%>.AGES.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.AGES.VALUE)=FALSE OR INSTR(1,<%=SELF%>.AGES.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.AGES.VALUE=""
			<%=SELF%>.AGES.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
ELSEIF N=5 THEN
	IF TRIM(<%=SELF%>.GTDAT.VALUE)<>"" THEN
		IF ISNUMERIC(<%=SELF%>.GTDAT.VALUE)=FALSE OR INSTR(1,<%=SELF%>.GTDAT.VALUE,"-")>0 THEN
			ALERT "輸入錯誤!!請輸入正確數字"
			<%=SELF%>.GTDAT.VALUE=""
			<%=SELF%>.GTDAT.FOCUS()
			EXIT FUNCTION
		END IF
	END IF
END IF

END FUNCTION
 
</script>


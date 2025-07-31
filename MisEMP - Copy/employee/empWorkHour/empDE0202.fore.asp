<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!-- #include file="../../Include/func.inc" -->
<!--#include file="../../include/sideinfolev2.inc"-->
<%
SESSION.CODEPAGE="65001"
SELF = "empde0202"

'--------------------------------------------------------------------------------------
FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END IF
END FUNCTION
'--------------------------------------------------------------------------------------

gTotalPage = 1
PageRec = 15    'number of records per page
TableRec = 20    'number of fields per record    

dat1=request("dat1")
dat2=request("dat2") 
if dat1="" then dat1=year(date())&"/"&month(date())&"/01"
if dat2="" then dat2=date()
S_empid = request("S_empid") 
gid = request("gid")  

'response.write dat1 &"<BR>"
'response.write dat2 &"<BR>"
'response.end 

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset") 


sql="select b.whsno,  b.groupid, b.gstr, b.empnam_cn, b.empnam_vn, b.nindat, c.timeup as B_timeup , c.timedown as B_timedown , a.* from "&_
	"( select convert(char(10),dat, 111) as wdat , * from empforget where convert(char(10), dat, 111) between '"& dat1 &"' and '"& dat2 &"'  "&_
	"and isnull(status,'')<>'D'   ) a "&_
	"join ( select* from view_empfile  ) b on b.empid = a.empid  "&_
	"left join ( select* from empwork  ) c on c.empid = a.empid  and c.workdat = convert(char(8), a.dat ,112)  "&_
	"where  ( b.empid like '%"& S_empid &"'  or b.empnam_cn like '"& S_empid &"' ) and  b.groupid like '%"& gid &"' order by a.empid, a.dat " 
'response.write sql
	
'response.end  
if request("TotalPage") = "" or request("TotalPage") = "0" then
	CurrentPage = 1 	 	  	
	'sql="select b.whsno,  b.groupid, b.gstr, b.empnam_cn, b.empnam_vn, b.nindat, c.timeup as B_timeup , c.timedown as B_timedown , a.* from "&_
	'	"( select convert(char(10),dat, 111) as wdat , * from empforget where convert(char(10), dat, 111) between '"& dat1 &"' and '"& dat2 &"'  "&_
	'	"and isnull(status,'')<>'D'   ) a "&_
	'	"join ( select* from view_empfile  ) b on b.empid = a.empid  "&_
	'	"left join ( select* from empwork  ) c on c.empid = a.empid  and c.workdat = convert(char(8), a.dat ,112)  " 
	'response.write sql 
	rs.Open SQL, conn, 3, 3  
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
			tmpRec(i, j, 2) = trim(rs("lsempid"))
			tmpRec(i, j, 3) = trim(rs("wdat"))
			tmpRec(i, j, 4) = trim(rs("timeup"))
			tmpRec(i, j, 5) = trim(rs("timedown"))
			tmpRec(i, j, 6) = trim(rs("toth"))
			tmpRec(i, j, 7) = trim(rs("groupid"))		
			tmpRec(i, j, 8) = trim(rs("gstr"))			
			tmpRec(i, j, 9) = trim(rs("nindat"))
			tmpRec(i, j, 10) = trim(rs("empnam_cn"))
			tmpRec(i, j, 11) = trim(rs("empnam_vn"))
			tmpRec(i, j, 12) = trim(rs("whsno"))
			tmpRec(i, j, 13) = trim(rs("B_timeup"))
			tmpRec(i, j, 14) = trim(rs("B_timedown"))
			tmpRec(i, j, 15) = trim(rs("empnam_cn"))&trim(rs("empnam_vn"))
			tmpRec(i, j, 16) =rs("autoid")
			'tmpRec(i, j, 17) =rs("caB3")
			tmpRec(i, j, 17) ="" 'Steven 20200529 rs("caB3") not exist
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
	Session("empde0202") = tmpRec
else
	TotalPage = (request("TotalPage"))	
	gTotalPage = cint(request("gTotalPage"))	
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	'StoreToSession() 
	tmpRec = Session("empde0202") 	
	
	Select case request("send")
	     Case "FIRST"
		      CurrentPage = 1
	     Case "BACK"
		      if cint(CurrentPage) <> 1 then
			     CurrentPage = CurrentPage - 1
		      end if
	     Case "NEXT"
		      if cint(CurrentPage) <= cint(gTotalPage) then
			     CurrentPage = CurrentPage + 1
		      end if
	     Case "END"
		      CurrentPage = gTotalPage
	     Case Else
		      CurrentPage = 1
	end Select
end if

%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="refresh" >
<link rel="stylesheet" href="../../Include/style.css" type="text/css">
<link rel="stylesheet" href="../../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	<%=self%>.dat1.focus()
	<%=self%>.dat1.select()
end function

function colschg(index)
	'thiscols = document.activeElement.name
	'if window.event.keyCode = 38 then
	'	IF INDEX<>0 THEN
	'		document.all(thiscols)(index-1).SELECT()
	'	END IF
	'end if
	'if window.event.keyCode = 40 then
	'	document.all(thiscols)(index+1).SELECT()
	'end if
end function 

function resch()
	<%=self%>.totalpage.value="0" 
	<%=self%>.submit()
end function 

-->
</SCRIPT>
</head>
<body  topmargin="5" leftmargin="5"  marginwidth="0" marginheight="0"  onkeydown="enterto()"  ONLOAD="F()" >
<form name="<%=self%>"  method="post" action = "<%=self%>.fore.asp" >
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>">
<INPUT TYPE=hidden NAME=pagerec VALUE="<%=pagerec%>">
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

	<table border=0 width="100%">
		<tr>
			<td >
				<table class="txt" cellspacing="3" cellpadding="3">					
					<tr height=30>
						<td><a href="empDe0201.asp?pgid=<%=request("pgid")%>"><font color=blue>忘刷卡資料輸入NHẬP QUẢN LÝ QUÊN GẠT THẺ</font></a></td>
						<td width="50px">&nbsp;</td>
						<td>忘刷卡資料刪查KIỂM TRA QUẢN LÝ QUÊN GẠT THẺ</td>
					</tr>
				</table>
			</td>
		</tr>		
		<tr>
			<td>
				<fieldset style="margin:0;padding:0;width=90%"><legend><font class=txt9>資料查詢 KIỂM TRA DỮ LIỆU</font></legend>
				<table   class="txt" cellspacing="3" cellpadding="3">	
					<tr height=35>
						<td align=right>日期<br>Ngày tháng:</td>
						<td nowrap>
							<input type="text" style="width:120px" name=dat1  value="<%=fdt(dat1)%>"  onblur="date_change(1)">~
							<input type="text" style="width:120px" name=dat2  value="<%=fdt(dat2)%>" onblur="date_change(2)">	
						</td>
						<td  align=right nowrap>查詢內容<br>Nội dung:</td>
						<td>
							<input type="text" style="width:100px" name=S_empid  value="<%=S_empid%>">	 		
						</td>
						<td width=50 align=right>單位<br>Bộ phận:</td>
						<td>
							<select name=gid  onchange="resch()" style="width:120px">
								<option value="">全部 ALL</option>
								<%SQL="SELECT * FROM BASICCODE WHERE FUNC='GROUPID' and sys_type <>'AAA' ORDER BY SYS_TYPE "
									SET RST = CONN.EXECUTE(SQL)
									WHILE NOT RST.EOF
								%>
								<option value="<%=RST("SYS_TYPE")%>" <%IF  RST("SYS_TYPE")=gid THEN %> SELECTED <%END IF%> ><%=RST("SYS_VALUE")%></option>
								<%
									RST.MOVENEXT
									WEND
									set rst=nothing
								%>
							</select>
						</td>
						<td align=right>
							<input type=button name=send value="查詢" class="btn btn-sm btn-outline-secondary" onclick="resch()"  onkeydown="resch()" >
						</td>
					</tr>	
				</table>
				</fieldset>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="98%">	
					<tr height=30 bgcolor="#EAEAEA">
						<td align=center>XÓA</td>
					<td align=center>員工編號<br>Mã số nhân viên</td>
					<td align=center>員工姓名<br>Họ tên nhân viên</td>
					<td align=center>單位 <br>Bộ phận</td>
					<td align=center>臨時卡號<br>Số thẻ tạm thời</td>
					<td align=center>出勤日期<br>Ngày công tác</td>
					<td align=center>(夜班)<br>Ca đêm</td>
					<td align=center>上班<br>Lên ca</td>
					<td align=center>下班<br>Xuống ca</td>
					<td align=center>時數<br>Số giờ</td>
					<td align=center bgcolor="#FFECF8">原上班<br>Nguyên lên ca</td>
					<td align=center bgcolor="#FFECF8">原下班<br>Nguyên xuống ca</td>
					</tr> 
					<%
						for CurrentRow = 1 to PageRec			
							IF CurrentRow MOD 2 = 0 THEN
								WKCOLOR="LavenderBlush"
							ELSE
								WKCOLOR="#D9ECFF"
							END IF			
					%>
					<tr bgcolor="<%=wkcolor%>" height=25>	
						<td align="center">
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type=button name=btn value="DEL" class="btn btn-sm btn-outline-secondary" onclick="delchg(<%=currentRow-1%>)">
							<%else%>-
								<input type=hidden name=btn  >
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=empid class='readonly8' readonly8  value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>"  maxlength=5 >
								<input type=hidden  name=whsno class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" >
								<input type=hidden  name=autoid class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" >
							<%else%>
								<input type=hidden  name=empid >
								<input type=hidden  name=whsno >
								<input type=hidden  name=autoid >
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=empname class='readonly8' readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 15)%>">
							<%else%>
								<input type=hidden  name=empname >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=gstr class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 8)%>" >
								<input type=hidden name=groupid class='readonly' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" size=7>
							<%else%>
								<input type=hidden  name=gstr >
								<input type=hidden  name=groupid >
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=Lempid class='readonly8'  readonly value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>"  >
							<%else%>
								<input type=hidden  name=Lempid >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=wdat class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>"  maxlength=10   >
							<%else%>
								<input type=hidden  name=wdat >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%" name=caB3 class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 17)%>"  maxlength=1   >
							<%else%>
								<input type=hidden  name=caB3 >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text"  name=timeup class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>"  maxlength=5  style='width:100%;text-align:center'>
							<%else%>
								<input type=hidden  name=timeup >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" name=timedown class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 5)%>" maxlength=5  onblur="t2chg(<%=currentrow-1%>)" style='width:100%;text-align:center'>
							<%else%>
								<input type=hidden  name=timedown >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" name=toth class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>"  style='width:100%;text-align:right'>
							<%else%>
								<input type=hidden  name=toth >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%;text-align:center" name=B_timeup class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" >
							<%else%>
								<input type=hidden  name=B_timeup >			
							<%end if%>
						</td>
						<td>
							<%if tmpRec(CurrentPage, CurrentRow, 1)<>"" then %>	
								<input type="text" style="width:100%;text-align:center" name=B_timedown class='readonly8' readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 14)%>">
							<%else%>
								<input type=hidden  name=B_timedown >			
							<%end if%>
						</td>
					</tr>
					<%next%>  
					<input type=hidden name=btn  > 	
					<input type=hidden  name=empid >
					<input type=hidden  name=whsno >
					<input type=hidden  name=autoid >
					<input type=hidden  name=empname >
					<input type=hidden  name=gstr >
					<input type=hidden  name=groupid >
					<input type=hidden  name=Lempid >
					<input type=hidden  name=wdat >
					<input type=hidden  name=timeup >
					<input type=hidden  name=timedown >	
					<input type=hidden  name=toth >	
					<input type=hidden  name=B_timeup > 			
					<input type=hidden  name=B_timedown >
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<TABLE  class="txt" cellspacing="0" cellpadding="0">
					<tr>
						<td align="CENTER" height=40 width=70%>

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
						PAGE:<%=CURRENTPAGE%> / <%=GTOTALPAGE%> , COUNT:<%=RECORDINDB%>
						</td>	
					</TR>
				</TABLE> 
			</td>
		</tr>
	</table>	
</form>


</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("empde0202")
	for CurrentRow = 1 to PageRec		
		tmpRec(CurrentPage, CurrentRow, 1) = request("empid")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("lempid")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("wdat")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 4) = request("timeup")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 5) = request("timedown")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 6) = request("toth")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 7) = request("groupid")(CurrentRow)				
		tmpRec(CurrentPage, CurrentRow, 8) = request("gstr")(CurrentRow)						
		tmpRec(CurrentPage, CurrentRow, 12) = request("whsno")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 13) = request("b_timeup")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 14) = request("b_timedown")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 15) = request("empname")(CurrentRow)		
	next
	Session("empde0202") = tmpRec
End Sub
%> 
<script language=vbscript >
function empidchg(index)	
	if <%=self%>.empid(index).value<>"" then 
		codestr1= Ucase(trim(<%=self%>.empid(index).value)) 
			
		open "<%=SELF%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 "&code=" &	codestr1 , "Back"
		'PARENT.BEST.COLS="70%,30%"
		'DATACHG(INDEX) 
	end if  	
end  function 

function delchg(index)
	if confirm("Delete(Cancel) This Record?",64) then 
		empid = <%=self%>.empid(index).value 
		autoid = <%=self%>.autoid(index).value 		
		'open "<%=self%>.deldb.asp?code1="& autoid &"&code2="& empid , "Back"
		'parent.best.cols="70%,30%"
		<%=self%>.action ="<%=self%>.deldb.asp?code1="& autoid &"&code2=" & empid
		<%=self%>.submit()
	end if 	
end function 
 

'*******檢查日期 Kiểm tra ngày tháng*********************************************
FUNCTION date_change(a)

if a=1 then
	INcardat = Trim(<%=self%>.dat1.value)
elseif a=2 then
	INcardat = Trim(<%=self%>.dat2.value)
elseif a=3 then
	INcardat = Trim(<%=self%>.dat1.value)
elseif a=4 then
	INcardat = Trim(<%=self%>.dat2.value)
end if

IF INcardat<>"" THEN
	ANS=validDate(INcardat)
	IF ANS <> "" THEN
		if a=1 then
			Document.<%=self%>.dat1.value=ANS
		elseif a=2 then
			Document.<%=self%>.dat2.value=ANS
		elseif a=3 then
			Document.<%=self%>.dat1.value=ANS
		elseif a=4 then
			Document.<%=self%>.dat2.value=ANS
		end if
	ELSE
		ALERT "EZ0067:輸入日期不合法 !!"
		if a=1 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=2 then
			Document.<%=self%>.dat2.value=""
			Document.<%=self%>.dat2.focus()
		elseif a=3 then
			Document.<%=self%>.dat1.value=""
			Document.<%=self%>.dat1.focus()
		elseif a=4 then
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




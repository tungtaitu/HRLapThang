<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../../GetSQLServerConnection.fun" -->
<!-- #include file="../../ADOINC.inc" -->
<!--#include file="../../include/sideinfolev2.inc"-->
<%
SESSION.CODEPAGE="65001"
SELF = "empde0201"

'--------------------------------------------------------------------------------------
FUNCTION FDT(D)
IF D <> "" THEN
	Response.Write YEAR(D)&"/"&RIGHT("00"&MONTH(D),2)&"/"&RIGHT("00"&DAY(D),2)
END IF
END FUNCTION
'--------------------------------------------------------------------------------------

gTotalPage = 5
PageRec = 15    'number of records per page
TableRec = 20    'number of fields per record    


Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset") 

sql="select b.whsno,  b.groupid, b.gstr, b.empnam_cn, b.empnam_vn, b.nindat, c.timeup as B_timeup , c.timedown as B_timedown , a.* from "&_
	"( select convert(char(10),dat, 111) as wdat , * from empforget where convert(char(10), dat, 111)='2001/01/01' and isnull(status,'')<>'D'and isnull(whsno,'')='LA'  ) a "&_
	"join ( select* from view_empfile  ) b on b.empid = a.empid  "&_
	"left join ( select* from empwork  ) c on c.empid = a.empid  and c.workdat = convert(char(8), a.dat ,112)  " 
	
 
if request("gTotalPage") = "" or request("gTotalPage") = "0" then
	CurrentPage = 1
	rs.Open SQL, conn, 3, 3 
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array

	for i = 1 to gTotalPage
	 for j = 1 to PageRec
		if not rs.EOF then	
			tmpRec(i, j, 0) = "no"
			tmpRec(i, j, 1) = trim(rs("empid"))
			tmpRec(i, j, 2) = trim(rs("lempid"))
			tmpRec(i, j, 3) = trim(rs("wdat"))
			tmpRec(i, j, 4) = trim(rs("timeup"))
			tmpRec(i, j, 5) = trim(rs("timedown"))
			tmpRec(i, j, 6) = trim(rs("toth"))
			tmpRec(i, j, 7) = trim(rs("groupid"))		
			tmpRec(i, j, 8) = trim(rs("gstr"))			
			tmpRec(i, j, 9) = trim(rs("nindat"))
			tmpRec(i, j, 10) = trim(rs("empname_cn"))
			tmpRec(i, j, 11) = trim(rs("empname_vn"))
			tmpRec(i, j, 12) = trim(rs("whsno"))
			tmpRec(i, j, 13) = trim(rs("B_timeup"))
			tmpRec(i, j, 14) = trim(rs("B_timedown"))
			tmpRec(i, j, 15) = trim(rs("empname_cn"))&trim(rs("empname_vn"))&"-"&trim(rs("nindat"))
			tmpRec(i, j, 16) = trim(rs("caB3")) '是否夜班
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
	Session("empde0201") = tmpRec
else
	TotalPage = (request("TotalPage"))	
	gTotalPage = cint(request("gTotalPage"))	
	CurrentPage = cint(request("CurrentPage"))
	RecordInDB  = REQUEST("RecordInDB")
	
	tmpRec = Session("empde0201")
	StoreToSession() 
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
	<%=self%>.empid(0).SELECT()
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

<table border=0 style="width:100%">	
	<tr>
		<td>
			<table class="txt"  cellspacing="3" cellpadding="3" >
				<tr>					
					<td>忘刷卡資料輸入 NHẬP QUẢN LÝ QUÊN GẠT THẺ</td><td width="50px">&nbsp;</td>
					<td><a href="empDe0202.asp?pgid=<%=request("pgid")%>"><font color=blue>忘刷卡資料刪查 KIỂM TRA QUẢN LÝ QUÊN GẠT THẺ</font></a></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center">
			<table id="myTableGrid" width="98%">
				<tr height=30 bgcolor="#EAEAEA">
					<td align=center>STT</td>
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
					<td align='center'><%=(currentpage-1)*pagerec+CurrentRow%></td>
					<td> 
						<input type="text" style="width:100%" name=empid class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 1)%>"  maxlength=5 onblur="empidchg(<%=currentrow-1%>)">
						<input type=hidden  name=whsno class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 12)%>" >
					</td>
					<td><input type="text" style="width:100%" name=empname class='readonly8' readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 15)%>" ></td>
					<td>
						<input  type="text" style="width:70px" name=gstr class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 8)%>">
						<input type=hidden name=groupid class='readonly' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 7)%>" >
					</td>
					<td><input type="text" style="width:100%" name=Lempid class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 2)%>"  onblur="lempidchg(<%=currentrow-1%>)"></td>
					<td><input type="text" style="width:100%" name=wdat class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 3)%>"  maxlength=10 onblur="datchg(<%=currentrow-1%>)" ></td>
					<td align='center'>
					<input type="checkbox" name="fnisb3" <%if tmpRec(CurrentPage, CurrentRow, 16)="Y" then%>checked<%end if%> onclick="fnb3chg(<%=currentrow-1%>)" >
					<input type='hidden' name=cab3  class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 16)%>" size=2 maxlength=10  >
					</td>
					<td><input type="text"  name=timeup class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 4)%>"  maxlength=5 onblur="t1chg(<%=currentrow-1%>)" style='width:100%;text-align:center'></td>
					<td><input type="text"  name=timedown class=inputbox value="<%=tmpRec(CurrentPage, CurrentRow, 5)%>"   maxlength=5  onblur="t2chg(<%=currentrow-1%>)" style='width:100%;text-align:center'></td>
					<td><input type="text"  name=toth class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 6)%>"  style='width:100%;text-align:right'></td>
					<td><input type="text" style="width:100%" name=B_timeup class='readonly8' readonly value="<%=tmpRec(CurrentPage, CurrentRow, 13)%>" ></td>
					<td><input type="text" style="width:100%" name=B_timedown class='readonly8' readonly  value="<%=tmpRec(CurrentPage, CurrentRow, 14)%>" ></td>
				</tr>
				<%next%> 
			</table>
		</td>
	</tr>
	<tr>
		<td align="center">
			<TABLE class="txt"  cellspacing="3" cellpadding="3">
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
					<td>
						<input type=button name=send  value="Confirm" class="btn btn-sm btn-danger" onclick="go()" >
						<input type=reset name=send  value="Cancel" class="btn btn-sm btn-outline-secondary">
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
	tmpRec = Session("empde0201")
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
	Session("empde0201") = tmpRec

End Sub
%> 
<script language=vbscript >
function fnb3chg(index) 
	if <%=self%>.fnisb3(index).checked then 
		<%=self%>.caB3(index).value="Y"  
	else 	
		<%=self%>.caB3(index).value=""  
	end if 	
	lempidchg(index)
end function 
function empidchg(index)	
	if <%=self%>.empid(index).value<>"" then 
		codestr1= Ucase(trim(<%=self%>.empid(index).value)) 
			
		open "<%=SELF%>.back.asp?ftype=A&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 "&code=" &	codestr1 , "Back"
		'PARENT.BEST.COLS="70%,30%"
		'DATACHG(INDEX) 
	end if  	
end  function 	

function t1chg(index) 	
	if <%=self%>.timeup(index).value<>"" then 
		if isnumeric(left(<%=self%>.timeup(index).value,2))=false or isnumeric(right(<%=self%>.timeup(index).value,2))=false then 
			alert "時間格式輸入錯誤 nhập sai thời gian!!"
			<%=self%>.timeup(index).value=""
			<%=self%>.timeup(index).focus()
			exit function  
		end if 
		if left(<%=self%>.timeup(index).value,2)>="24" or right(<%=self%>.timeup(index).value,2)>60 then 
			alert "時間格式輸入錯誤 nhập sai thời gian!!"
			<%=self%>.timeup(index).value=""
			<%=self%>.timeup(index).focus()
			exit function  			
		end if 	 		
		<%=self%>.timeup(index).value  = left(<%=self%>.timeup(index).value,2)&":"& right(<%=self%>.timeup(index).value,2)
	end if 
end function   	

function t2chg(index)
	if <%=self%>.timedown(index).value<>"" then 
		if isnumeric(left(<%=self%>.timedown(index).value,2))=false or isnumeric(right(<%=self%>.timedown(index).value,2))=false then 
			alert "時間格式輸入錯誤 nhập sai thời gian!!"
			<%=self%>.timedown(index).value=""
			<%=self%>.timedown(index).focus()
			exit function  
		end if 
		if left(<%=self%>.timedown(index).value,2)>="24" or right(<%=self%>.timedown(index).value,2)>60 then 
			alert "時間格式輸入錯誤 nhập sai thời gian!!"
			<%=self%>.timedown(index).value=""
			<%=self%>.timedown(index).focus()
			exit function  			
		end if 	 		
		<%=self%>.timedown(index).value  = left(<%=self%>.timedown(index).value,2)&":"& right(<%=self%>.timedown(index).value,2)
		if <%=self%>.wdat(index).value <>"" then 
			if trim(<%=self%>.timeup(index).value)<>"" and trim(<%=self%>.timedown(index).value)<>"" then 
				DD1=<%=self%>.wdat(index).value&" "& trim(<%=self%>.timeup(index).value) 
				DD2=<%=self%>.wdat(index).value&" "& trim(<%=self%>.timedown(index).value) 
			end if  
			IF LEFT(<%=SELF%>.TIMEDOWN(INDEX).VALUE,2)< LEFT(<%=SELF%>.TIMEUP(INDEX).VALUE,2) THEN
				TOTH=ROUND( DATEDIFF("N", DD1, DD2 ) /30 ,0 ) /2 + 24
			ELSE
				TOTH=ROUND( DATEDIFF("N", DD1, DD2 ) /30 ,0 ) /2
			END IF
			<%=self%>.toth(index).value = TOTH 
			
			code1 = <%=self%>.empid(index).value
			code2 = <%=self%>.whsno(index).value
			code3 = <%=self%>.lempid(index).value
			code4 = <%=self%>.wdat(index).value
			code5 = <%=self%>.timeup(index).value
			code6 = <%=self%>.timedown(index).value
			code7 = <%=self%>.toth(index).value
			code8 = <%=self%>.cab3(index).value
			
			open "<%=SELF%>.back.asp?ftype=C&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 	 "&CODESTR01="& code1 &"&CODESTR02="& code2 &_
			 	 "&CODESTR03="& code3 &"&CODESTR04="& code4 &_
			 	 "&CODESTR05="& code5 &"&CODESTR06="& code6 &_
			 	 "&CODESTR07="& code7 , "Back"
			lempidchg(index)
			'PARENT.best.COLS="70%,30%"
		end if 	
	end if  
end  function   

function go()
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function  

function lempidchg(index)
	'if <%=self%>.lempid(index).value<>"" then 
		<%=self%>.lempid(index).value = Ucase(<%=self%>.lempid(index).value) 		
		code1 = <%=self%>.empid(index).value
		code2 = <%=self%>.whsno(index).value
		code3 = <%=self%>.lempid(index).value
		code4 = <%=self%>.wdat(index).value
		code5 = <%=self%>.timeup(index).value
		code6 = <%=self%>.timedown(index).value
		code7 = <%=self%>.toth(index).value
		code8 = <%=self%>.cab3(index).value
		
		open "<%=SELF%>.back.asp?ftype=C&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
		 	 "&CODESTR01="& code1 &"&CODESTR02="& code2 &_
		 	 "&CODESTR03="& code3 &"&CODESTR04="& code4  &_
		 	 "&CODESTR05="& code5 &"&CODESTR06="& code6 &_
		 	 "&CODESTR07="& code7 &"&CODESTR08="& code8, "Back"
	'PARENT.best.COLS="70%,30%"
	'end if 
end function 

function datchg(index)
	if <%=self%>.wdat(index).value<>"" then 
		INcardat = Trim(<%=self%>.wdat(index).value) 		
		IF INcardat<>"" THEN
			ANS=validDate(INcardat)
			IF ANS <> "" THEN				
				Document.<%=self%>.wdat(index).value=ANS 
				if trim(<%=self%>.empid(index).value)<>"" then 
					codestr1= Ucase(trim(<%=self%>.empid(index).value)) 
					codestr2= Ucase(trim(<%=self%>.wdat(index).value)) 
					codestr3= Ucase(trim(<%=self%>.cab3(index).value)) 
					open "<%=SELF%>.back.asp?ftype=B&index="& index &"&CurrentPage="& <%=CurrentPage%> & _
			 			 "&CODESTR01="& codestr1 & "&CODESTR02=" & codestr2& "&CODESTR03=" & codestr3 , "Back"
					'PARENT.best.COLS="70%,30%"
				end if 				
			ELSE
				ALERT "EZ0067:輸入日期不合法 NHẬP SAI !!"	
				Document.<%=self%>.wdat(index).value=""
				Document.<%=self%>.wdat(index).focus()
			end if 
		end if 		
	end if 
end  function   

'*******檢查日期 Kiểm tra ngày tháng****************************** 
'_________________DATE CHECK_____________________________________ 
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
'_________________________________________________________________ 
 



</script>



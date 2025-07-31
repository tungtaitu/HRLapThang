<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<META HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->
<!--#include file="../include/sideinfo.inc"-->
<%

if session("netuser")="" then 
	response.write "使用者帳號為空!!請重新登入!!"
	'response.end 
end if 	 

SELF = "YEBE0103C"


nowmonth = year(date())&right("00"&month(date()),2)
if month(date())="01" then
	calcmonth = year(date()-1)&"12"
else
	calcmonth =  year(date())&right("00"&month(date())-1,2)
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
<body   onkeydown="enterto()" onload="f()">
<form name="<%=self%>" method="post" action="<%=self%>.upd.asp">
<INPUT TYPE=HIDDEN NAME="UID" VALUE=<%=SESSION("NETUSER")%>>
<INPUT TYPE=hidden NAME=nowmonth VALUE="<%=nowmonth%>">
 
	<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="80%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table  cellspacing="1" cellpadding="1" class="table table-bordered table-sm bg-white text-secondary" bgcolor=black>
								<tr bgcolor=#ffffff height=35>
									<Td align=center bgcolor="#ffffff" width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand" ><a href="yebe0103.asp">外包資料新增<br>tu lieu moi thau ngoai</a></td>
									<Td align=center bgcolor="#ffffff"  width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0104.fore.asp">外包資料維護<br>xoa/sua tu lieu thau ngoai</a></td>
									<Td align=center bgcolor="#ffffff"  width=160 onMouseOver="this.style.backgroundColor='#FFEB78'" onMouseOut="this.style.backgroundColor='#ffffff'"   style="cursor:hand"><a href="yebe0105.fore.asp">外包資料查詢<br>K.Tra tu lieu</a></td>
									<Td align=center bgcolor="#ffff99"  width=160  ><a href="yebe0103C.asp">照片上傳與新增<br>Update Photos</a></td>
								</tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<TABLE  CLASS=txt BORDER=0 cellspacing="2" cellpadding="2" >
								<td><td><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
								<TR >		 		
									<td><a href="vbscript:upwbphotos()"><font color=blue>傳照片(UpDate Photos)</a></font></td>
								</TR> 
								<td><td><hr size=0	style='border: 1px dotted #999999;' align=left ></td></tr>
							</table>
						</td>
					</tr>
					<tr>
						<td>
							<TABLE class="table table-bordered table-sm bg-white text-secondary">
							<%Set conn = GetSQLServerConnection()
							sql ="select aid,isnull(wbempid,'') wbempid, isnull(wbwhsno,'') wbwhsno, isnull(filename,'') filename  from wbphotos where isnull(wbempid,'')='' and isnull(filename,'')<>'' and isnull(status,'')<>'D' "
							Set rs = Server.CreateObject("ADODB.Recordset")  
							rs.open sql, conn, 3, 3
							for i = 1 to rs.recordcount 
							%>
								<%if i mod 5 = 1 then%><TR ><%end if%>
									<td width=110 nowrap>
										<div style='cursor:hand' onclick="editData(<%=i-1%>)">
										<img src="wbphotos/<%=rs("filename")%>" width=110 height=130 border=1 ><BR><font color=#993399><%=rs("filename")%></font>
										<input name=aid class="form-control form-control-sm mb-2 mt-2" type=hidden value="<%=rs("aid")%>">
										</div>
										<BR>Xuong:<%if rs("wbempid")="" then%><Select name=whsno class="form-control form-control-sm mb-2 mt-2" style='width:60'>
											<option value="LA">LA</option>
											<option value="DN">DN</option>
											<option value="BC">BC</option>
										</select><%else%><input name=whsno class="form-control form-control-sm mb-2 mt-2" type=hidden><%end if%>
										<BR>SO The:<%if rs("wbempid")="" then%><input name=wbid class="form-control form-control-sm mb-2 mt-2" size=6><%else%><input type=hidden  name=wbid class=inputbox size=6 value="<%=rs("")%>"><%end if%>			
										<hr size=0	style='border: 1px dotted #999999;' align=left>			
									</td>
								<%if i mod 5 = 0 then%></TR><%end if%>  
							<%
							rs.movenext
							next
							rs.close
							conn.close
							set conn=nothing
							set rs=nothing 
							%>		
							<input name=aid class="form-control form-control-sm mb-2 mt-2" type=hidden>
							<input name=whsno class="form-control form-control-sm mb-2 mt-2" type=hidden>
							<input name=wbid class="form-control form-control-sm mb-2 mt-2" type=hidden>
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
<script language=vbscript>
 
'-----------------enter to next field 
function getlorry()
	open "getlorry.asp", "Back"
	parent.best.cols="50%,50%"
end function  

function editData(index) 
	x=<%=self%>.aid(index).value
	wbid=<%=self%>.wbid(index).value
	whsno=<%=self%>.whsno(index).value 	
	if wbid<>"" then 
		open "YEBE0104.foregnd.asp?wbid="& wbid &"&wbphotoid="& x  , "_balnk", "top=20, left=20, width=600, height=600, scrollbars=yes"
	else
		open "<%=self%>.index.asp?wbphotoid=" & x , "_balnk", "top=20, left=20, width=600, height=600, scrollbars=yes"
	end if 	
end function 

function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9
end function

function f()
	'<%=self%>.whsno.focus()
	'<%=self%>.EMPID.select()
end function

function groupchg()
	code = <%=self%>.GROUPID.value
	open "<%=self%>.back.asp?ftype=groupchg&code="&code , "Back"
	'parent.best.cols="50%,50%"
end function

function unitchg()
	code = <%=self%>.unitno.value
	open "<%=self%>.back.asp?ftype=UNITCHG&code="&code , "Back"	
	'parent.best.cols="50%,50%"
end function 

function upwbphotos()
	open "send.muchfile.asp" , "_blank" , "left=20 top=20 width=600 scrollbars=yes"
end function

function whsnochg()	
	code1 = <%=self%>.whsno.value
	code2 = <%=self%>.wbloai.value
	if code1<>"" and code2<>"" then 		
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"			
		parent.best.cols="100%,0%"
	end if 
end function  

function loaichg()
	code1 = <%=self%>.wbloai.value
	code2 = <%=self%>.whsno.value
	if code1<>"" and code2<>"" then 
		open "<%=self%>.back.asp?ftype=getwbid&code1="&code1 &"&code2="& code2  , "Back"	
		'parent.best.cols="50%,50%"
	end if
end function 

function empidchg()
	empidstr = Ucase(Trim(<%=self%>.empid.value))
	if empidstr<>"" then
		open "<%=self%>.back.asp?ftype=empidchk&code="& empidstr , "Back"
		'parent.best.cols="50%,50%"
	end if
end function

function sexchg(x)
	if <%=self%>.radio1(0).checked=true then
		<%=self%>.sexstr.value="M"
	elseif 	<%=self%>.radio1(1).checked=true then
		<%=self%>.sexstr.value="F"
	else
		<%=self%>.sexstr.value=""
	end if
end function

function marrychg(x)
	if <%=self%>.radio2(0).checked=true then
		<%=self%>.marryed.value="Y"
	elseif 	<%=self%>.radio2(1).checked=true then
		<%=self%>.marryed.value="N"
	elseif 	<%=self%>.radio2(2).checked=true then
		<%=self%>.marryed.value="L"
	else
		<%=self%>.marryed.value=""	
	end if
end function

function BACKMAIN()
	open "../main.asp" , "_self"
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


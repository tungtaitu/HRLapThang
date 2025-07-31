<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%>
<!-- #include file = "../GetSQLServerConnection.fun" -->
<!-- #include file="../ADOINC.inc" -->

<%
Set conn = GetSQLServerConnection()
self="yeee03" 

if  instr(conn,"168")>0 then 
	w1="LA"
elseif  instr(conn,"169")>0 then 
	w1="DN"	
elseif  instr(conn,"47")>0 then 
	w1="BC"	
end if 	 

w1=session("mywhsno") 

sql="select njym, convert(char(10),td1,111) td1 ,convert(char(10),td2,111) td2,convert(char(20),mdtm,120) mdt , muser from empNJYM_set order by njym desc  " 
set rs=conn.execute(Sql) 
i=1
year1= cstr(year(date())-1) 
'response.write year1 
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="refresh">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">

<link rel="stylesheet" type="text/css" href="../template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="../template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="../template/css/mis.css">
<link rel="stylesheet" type="text/css" href="../template/datepicker/datepicker.css">
</head>
<body   onkeydown="enterto()"   >
<form name="<%=self%>" method="post" >
	<table border="0" cellspacing="0" cellpadding="0">
		<tr>
			<TD><img border="0" src="../image/icon.gif" align="absmiddle">年假統計設定</TD>
		</tr>
	</table>
	<hr size=0	style='border: 1px dotted #999999;' align=left>
	<table width="100%" border=0 >
		<tr>
			<td>
				<table width="98%" BORDER=0 align=center cellpadding=0 cellspacing=0 >
					<tr>
						<td>
							<table class="table table-bordered table-sm bg-white text-secondary">
								<TR height=30 bgcolor="#cccccc" class="txt8">				
									<td nowrap align=center >年度</td>
									<td nowrap align=center >統計日期<br>(起)</td>
									<td nowrap align=center >統計日期<br>迄</td>
									<td nowrap align=center >異動日期<br>Change date</td>
									<td nowrap align=center >修改<br>User</td>
								</TR>
								<Tr  class="txt8">
										<td><input name="njym" class="inputbox" size=5 value="" maxlength=6><input name=flags value="U" size=1 class="inputbox" type="hidden"></td>
										<td><input name="Td1" class="inputbox" id="td1" size=12 value="" onblur="datchg(0,'td1','OTD1')"></td>
										<td><input name="Td2" class="inputbox" id="td2" size=12 value="" onblur="datchg(0,'td2','OTD2')"></td>
										<td>&nbsp;<input name="mdt" value="" type="hidden" ></td>
										<td>&nbsp;
										<input name="oTd1" class="inputbox" size=12 id="otd1" value="" type="hidden">
										<input name="oTd2" class="inputbox" size=12 id="otd2" value="" type="hidden">
										</td>
										
								</tr>		
								<%if not rs.eof then%>	
								<%while not rs.eof 
									i = i + 1 
									if rs("njym") < year1 then 
										modes="readonly" 
										clstyp="readonly"				
									else 
										modes="" 
										clstyp="inputbox"
									end if	
								%>
									<Tr  class="txt8">
										<td><input name="njym" class="readonly" readonly size=5 value="<%=rs("njym")%>">
										<input name=flags value="" size=1 class="inputbox" type="hidden" >
										</td>
										<td><input name="Td1" class="<%=clstyp%>" <%=modes%> size=12 id="td1" value="<%=rs("td1")%>" onblur="datchg(<%=i%>,'td1','OTD1')"></td>
										<td><input name="Td2" class="<%=clstyp%>" <%=modes%>  size=12 id="td2" value="<%=rs("td2")%>" onblur="datchg(<%=i%>,'td2','OTD2')" ></td>
										<td><%=rs("mdt")%><input name="mdt" value="<%=rs("mdt")%>" type="hidden" ></td>
										<td><%=rs("muser")%>
										<input name="oTd1" class="inputbox" size=12 id="otd1" value="<%=rs("td1")%>" type="hidden">
										<input name="oTd2" class="inputbox" size=12 id="otd2" value="<%=rs("td2")%>" type="hidden">
										</td>
									</tr>
								<%rs.movenext%>
								<%wend%>
								<%end if%>
								
							</table>
						</td>
					</tr>
					<tr>	
						<td align="center">
							<input name=flags value="" size=1 class="inputbox" type="hidden" >
							<input name="njym" value="" type="hidden" >
							<input name="td1" value="" type="hidden" >
							<input name="td2" value="" type="hidden" >
							<input name="mdt" value="" type="hidden" >
							<input name="oTd1" class="inputbox" id="otd1" size=12 value="" type="hidden">
							<input name="oTd2" class="inputbox" id="otd2" size=12 value="" type="hidden">
							<table class="table-borderless table-sm bg-white text-secondary">
								<tr>
									<td align=center>
										<input type=button  name=btm class="btn btn-sm btn-outline-secondary" value="確   認" onclick="go()" onkeydown="go()">
										<input type=reset  name=btrs class="btn btn-sm btn-outline-secondary" value="取   消">				
										<input name="pagerec" class=readonly8 readonly value="<%=i%>" size=3 >
									</td>
								</tr>
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
<!-- #include file="../Include/func.inc" -->

<script language=vbs>
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

function strchg(a)
	if a=1 then
		<%=self%>.empid1.value = Ucase(<%=self%>.empid1.value)
	elseif a=2 then
		<%=self%>.empid2.value = Ucase(<%=self%>.empid2.value)
	end if
end function

function go()	
	if <%=self%>.td1(0).value<>"" and <%=self%>.td2(0).value<>"" then 
		if <%=self%>.njym(0).value="" then 
			alert "請輸入年度(NAM)"
			<%=self%>.njym(0).focus()
			exit function
		end if 	
	end if 
 	<%=self%>.action="<%=self%>.backupd.asp"
 	<%=self%>.submit()
end function

function goexcel()
	if <%=self%>.yymm.value="" then 
		alert "請輸入[統計年度]!!"
		<%=self%>.yymm.focus()
		exit function 
	end if 	
	'open "<%=self%>.toexcel.asp" , "Back" 
	<%=self%>.action="<%=self%>.toexcel.asp"
	<%=self%>.target="Back"
	<%=self%>.submit()
	'parent.best.cols="50%,50%"
end function  

function datchg(index,sid,osid)
	ansstr = document.forms("<%=self%>").Elements(sid)(index).value
	if ansstr<>"" THEN
		ANS=validDate(ansstr)
		iF ANS <> "" THEN
			document.forms("<%=self%>").Elements(sid)(index).value = ans  		
			if <%=self%>.td1(index).value<><%=self%>.otd1(index).value or  <%=self%>.td2(index).value<><%=self%>.otd2(index).value then 
				<%=self%>.flags(index).value="U"
			else	
				<%=self%>.flags(index).value=""
			end if 
		ELSE
			ALERT "EZ0067:輸入日期不合法 !!"
			document.forms("<%=self%>").Elements(sid)(index).value=""
			document.forms("<%=self%>").Elements(sid)(index).focus()
		end if	
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
</script> 
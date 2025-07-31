<%@ Language=VBScript CODEPAGE=65001%>
<!---------  #include file="../../GetSQLServerConnection.fun"  -------->
<!-- #include file="../Include/SIDEINFO.inc" -->

<%
SELF = "YTBAE03"

Set conn = GetSQLServerConnection()
gTotalPage = 1
PageRec = 10    'number of records per page
TableRec = 20    'number of fields per record
proc1 = request("proc1")
if proc1="" then proc1="CM"
'sqlx="select b.* from ( select * from ydbmcode where tblid='s17'  )   a  "&_
'		 "left  join ( select * from ytbmproc ) b on b.loai = a.tblcd "&_
'		 "where b.proctype='"&proc1&"' "
'		 'response.write sqlx
'Set RSx = Server.CreateObject("ADODB.Recordset")
'rsx.open sqlx, conn, 1, 3
'if rs.eof then
'	sql="select '"&proc1&"' as proc1 , a.tblcd, a.tbldesc, b.* from ( select * from ydbmcode where tblid='s17'  )   a  "&_
'			"left  join ( select * from ytbmproc ) b on b.loai = a.tblcd "&_
'		 	"order by a.tblcd "
'else
	sql="select a.proc1 , a.tblcd, a.tbldesc,  b.* from "&_
			"( select '"&proc1&"' as proc1 , * from ydbmcode where tblid='s17'  )   a  "&_
			"left  join (select * from ytbmproc ) b on isnull(b.loai,'') = a.tblcd  and isnull(b.proctype,'') = a.proc1 "&_
		 	"where a.proc1='"&proc1&"' order by a.tblcd  "
'end if
'set rsx=nothing
'response.write sql
Set RS = Server.CreateObject("ADODB.Recordset")
rs.open sql, conn, 1, 3
if not rs.eof then
	PageRec =  rs.RecordCount
	rs.PageSize = PageRec
	RecordInDB = rs.RecordCount
	TotalPage = rs.PageCount
	gTotalPage = totalpage
end if
	CurrentPage = 1
	Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array
	for i = 1 to TotalPage
		for j = 1 to PageRec
			if not rs.EOF then
				tmpRec(i, j, 0) = "no"
				tmpRec(i, j, 1) = rs("aid")
				tmpRec(i, j, 2) = rs("proc1")
				tmpRec(i, j, 3) = rs("tblcd")
				tmpRec(i, j, 4) = rs("tbldesc")
				tmpRec(i, j, 5) = rs("zcq")
				tmpRec(i, j, 6) = rs("zcqname")
				tmpRec(i, j, 7) = rs("zcqmail")
				tmpRec(i, j, 8) = rs("hcq01")
				tmpRec(i, j, 9) = rs("hcq01name")
				tmpRec(i, j, 10) = rs("hcq01mail")
				tmpRec(i, j, 11) = rs("hcq02")
				tmpRec(i, j, 12) = rs("hcq02name")
				tmpRec(i, j, 13) = rs("hcq02mail")
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
	'Session("YTBDE02B") = tmpRec
 %>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=UTF-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../../ydb/Include/style.css" type="text/css">
<link rel="stylesheet" href="../../ydb/Include/style2.css" type="text/css">
</head>

<body background="bg_blue.gif"  topmargin=50 onkeydown=enterto() onload=f()>
<FORM method="POST" name="<%=self%>"  action="<%=self%>.fore.asp" >
<table width=600><tr><td  >
	<TABLE CELLPADDING=2 CELLSPACING=2 BORDER=0 >
	<TR>
			<TD align=right >簽核流程</TD>
		<TD>
		 <td width=400>
				<select name=proc1 class=txt onchange='submit()' >
					<option value="CM" <%if proc1="CM" then%>selected<%end if%>>(1)需料(請購)流程設計
					<option value="DDH" <%if proc1="DDH" then%>selected<%end if%>>(2)訂購(採購)流程設計
				</select>
		</TD>
	</tr>
	</TABLE>
	<table CELLPADDING=2 CELLSPACING=1 BORDER=0  bgcolor=#e4e4e4 class=txt>
		<tr bgcolor=#ffffff height=25>
			<td align=center>類別</td>
			<td align=center bgcolor="#FFCC99">審核主管</td>
			<td align=center bgcolor="#99CC99">核決主管(1)</td>
			<td align=center bgcolor="#FFFF99">核決主管(2)</td>
		</tr>
		<%for x = 1 to pagerec%>
		<tr bgcolor=#ffffff >
			<Td valign=top>
				<input type=hidden name=loai value="<%=tmprec(1,x,3)%>" size=5 class=readonly8 readonly  style='height:22px' >
				<input name=loai2 value="<%=tmprec(1,x,3)%>-<%=tmprec(1,x,4)%>" size=25 class=readonly8 readonly  style='height:22px'>
			</td>
			<Td nowrap>
				<input name=zcq value="<%=tmprec(1,x,5)%>" size=8 class=inputbox8  onchange="chkempid(<%=x-1%>)"  style='background-color:#ECF7FF;height:22px' ondblclick="gotcq(<%=x-1%>)" >
				<input name=zcqname value="<%=tmprec(1,x,6)%>" size=17 class=readonly8 readonly    style='height:22px'>
				<br><input name=zcqmail value="<%=tmprec(1,x,7)%>" size=28 class=inputbox8    style='height:22px'>
			</td>
 			<Td nowrap>
				<input name=hcq01 value="<%=tmprec(1,x,8)%>" size=8 class=inputbox8    onchange="chkempid(<%=x-1%>)"  style='background-color:#ECF7FF;height:22px' ondblclick="gotcq(<%=x-1%>)">
				<input name=hcq01name value="<%=tmprec(1,x,9)%>" size=17 class=readonly8 readonly    style='height:22px'>
				<br><input name=hcq01mail value="<%=tmprec(1,x,10)%>" size=28 class=inputbox8    style='height:22px'>
			</td>
			<Td nowrap>
				<input name=hcq02 value="<%=tmprec(1,x,11)%>" size=8 class=inputbox8  onchange="chkempid(<%=x-1%>)"    style='background-color:#ECF7FF;height:22px' ondblclick="gotcq(<%=x-1%>)">
				<input name=hcq02name value="<%=tmprec(1,x,12)%>" size=17 class=readonly8 readonly    style='height:22px'>
				<br><input name=hcq02mail value="<%=tmprec(1,x,13)%>" size=28 class=inputbox8    style='height:22px'>
			</td>
		</tr>
		<%next%>
	</table>
	<br>
	<table width="100%">
	<tr>
	 <td align="CENTER">
	 <input TYPE="button" name="send" VALUE="(Y)確認(Confirm)" class="button"  onclick=go() >
	 <input type="reset" name="send" value="(N)取消(Cancel)" class="button"   onclick="NEXTONE()">
	 <input name=pagerec value="<%=pagerec%>" type=hidden >

	 </td>
	</tr>
	</table>
</td></tr></table>
</FORM>

</BODY>
</HTML>


</script>
<script language="vbscript">
function f()
	<%=self%>.proc1.focus()
end function

function chkempid(index)
	set objall=document.all.<%=self%>
	thiscols=document.activeElement.name

	select case thiscols
			case "zcq"
				inti = 3+(index*11)
			case "hcq01"
				inti = 6+(index*11)
			case "hcq02"
				inti = 9+(index*11)
			case else
	end select
	if objall.item(inti).value="" then
		objall.item(inti+1).value=""
		objall.item(inti+2).value=""
		objall.item(inti).focus()
	else
		cols01=objall.item(inti).name
		cols02=objall.item(inti+1).name
		cols03=objall.item(inti+2).name
		code1 =objall.item(inti).value
		open "<%=self%>.back.asp?func=chkempid&code1="& code1 &"&cols01="& cols01 &_
			   "&cols02="& cols02 &"&cols03="& cols03 &"&index="& index , "Back"
				 'parent.best.cols="50%,50%"
	end if

	'for inti=0 to objall.length-1
	'	if objall.item(inti).value<>"" then
			'alert inti &" "& objall.item(inti).value
			'alert inti &" "& objall.item(inti).name
	'		select case objall.item(inti).name
	'		case "zcq", "hcq01" , "hcq02"
	'			cols01=objall.item(inti).name
	'			cols02=objall.item(inti+1).name
	'			cols03=objall.item(inti+2).name
	'			code1 =objall.item(inti).value
	'			open "<%=self%>.back.asp?func=chkempid&code1="& code1 &"&cols01="& cols01 &_
	'					 "&cols02="& cols02 &"&cols03="& cols03 &"&index="& index , "Back"
	'			parent.best.cols="50%,50%"
	'		case else
	'		end select
	'	end if
	'next

end function

function enterto()
 	if window.event.keyCode = 13 then
 		'取得表單元素名稱
 		thiscols=document.activeElement.name
 		if thiscols<>"memostr" and thiscols<>"spec" then
			 window.event.keyCode =9
		end if
	end if
	if window.event.keyCode = 113 then
		go()
	end if
end function

function NEXTONE()
	open "<%=self%>.asp", "_self"
end function

function gotcq(index)
	thiscols=document.activeElement.name
	open "getcqdata.asp?index="& index &"&cols1=" & thiscols ,"Back"
	parent.best.cols="50%,50%"
end function

function go()
	<%=self%>.action="<%=self%>.upd.asp"
	<%=self%>.submit()
end function

</script>



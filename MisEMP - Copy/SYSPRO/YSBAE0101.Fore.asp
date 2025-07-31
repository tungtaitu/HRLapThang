<%@Language=VBScript Codepage=65001 %>
<%Response.Buffer =True%>  
<!--#include file="../GetSQLServerConnection.fun"--> 
<%

const action	="YSBAE0101.FORE.ASP"
const formname	="FRM"
const method	="POST"

pid = request("pid")
upid = request("upid")
level = request("level")
vpath = request("vpath")
pname = request("pname")
pnameVN = request("pnameVN") 

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")     

gTotalPage = 50
PageRec = 10   'number of records per page
TableRec = 10    'number of fields per record
Query = TRIM(request("SearchKey"))

IF TRIM(QUERY)="" THEN  	
	if request("TotalPage") = "" or request("TotalPage") = "0" then 
		CurrentPage = 1	  	
	  	Source = "select * from SYSPROGRAM order by PROGRAM_ID"	 
	  	rs.Open Source, conn, 3, 3		
	  	IF NOT RS.EOF THEN 
			rs.PageSize = PageRec 
			RecordInDB = rs.RecordCount 
			TotalPage = rs.PageCount 
			gTotalPage = TotalPage + 1
		END IF 
			
		Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array		
		for i = 1 to TotalPage 
			for j = 1 to PageRec
				if not rs.EOF then 
					tmpRec(i, j, 0) = "no"
					tmpRec(i, j, 1) = trim(rs("PROGRAM_ID"))
					tmpRec(i, j, 2) = trim(rs("PROGRAM_NAME"))
					tmpRec(i, j, 3) = rs("LAYER_UP")
					tmpRec(i, j, 4) = rs("LAYER")
					tmpRec(i, j, 5) = rs("VIRTUAL_PATH")
					tmpRec(i, j, 6) = rs("PRONAME_VN") 
					rs.MoveNext 
				ELSE
					Set rs = nothing
					exit for 	
				end if 
				'if rs.EOF then 					
				'	Set rs = nothing
				'	exit for 
			 	'end if 	
			 next 		
			
		next 
		Session("SYSPRO01") = tmpRec	
	else    
	  	TotalPage = cint(request("TotalPage"))
	  	'response.write "TotalPage=" & TotalPage  &"<BR>"
	  	topage = cint(request("topage"))
		if topage="" then topage=0
		'response.write "topage=" & topage  &"<BR>"
	  	StoreToSession()
	  	CurrentPage = cint(request("CurrentPage"))
	  	RecordInDB = request("RecordInDB")
	  	tmpRec = Session("SYSPRO01")
	
	  	Select case request("send") 
		     Case "第一頁"
			      CurrentPage = 1			
		     Case "上一頁"
			      if cint(CurrentPage) <> 1 then 
				     CurrentPage = CurrentPage - 1				
			      end if
		     Case "下一頁"
			      if cint(CurrentPage) <= cint(gTotalPage) then 
				     CurrentPage = CurrentPage + 1 
			      end if			
		     Case "最末頁"
			      CurrentPage = TotalPage 			
		     Case Else 
			      CurrentPage = topage	
	  	end Select 
	end if
ELSE
	
		CurrentPage = 1	  	
	  	Source = "select * from SYSPROGRAM WHERE PROGRAM_ID LIKE '"& LEFT(QUERY,2) &"%' order by PROGRAM_ID"	 
	  	rs.Open Source, conn, 3, 3		
	  	IF NOT RS.EOF THEN 
			rs.PageSize = PageRec 
			RecordInDB = rs.RecordCount 
			TotalPage = rs.PageCount 
			gTotalPage = TotalPage + 1
		END IF 
			
		Redim tmpRec(gTotalPage, PageRec, TableRec)   'Array		
		for i = 1 to TotalPage 
			for j = 1 to PageRec
				if not rs.EOF then 
					tmpRec(i, j, 0) = "no"
					tmpRec(i, j, 1) = trim(rs("PROGRAM_ID"))
					tmpRec(i, j, 2) = trim(rs("PROGRAM_NAME"))
					tmpRec(i, j, 3) = rs("LAYER_UP")
					tmpRec(i, j, 4) = rs("LAYER")
					tmpRec(i, j, 5) = rs("VIRTUAL_PATH")
					tmpRec(i, j, 6) = rs("PRONAME_VN") 
					rs.MoveNext 
				end if 
			 next 		
			 if rs.EOF then 
				'rs.Close 
				Set rs = nothing
				exit for 
			 end if 
		next 
		Session("SYSPRO01") = tmpRec	
END IF 	

  
%>  

<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<link rel="stylesheet" href="../Include/style.css" type="text/css">
<link rel="stylesheet" href="../Include/style2.css" type="text/css">
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
<!--
function m(index)
   <%=SELF%>.send(index).style.backgroundcolor="lightyellow"
   <%=SELF%>.send(index).style.color="red"
end function

function n(index)
   <%=SELF%>.send(index).style.backgroundcolor="khaki"
   <%=SELF%>.send(index).style.color="black"
end function

function sortchg()
	<%=self%>.action="YDBQE0201B.ForeGnd.asp?totalpage=0"
	<%=self%>.submit
end function 
'-----------------enter to next field
function enterto()
	if window.event.keyCode = 13 then window.event.keyCode =9 		
end function

function gosrch()
	if trim(<%=formname%>.SearchKey.value)="" then 
		<%=formname%>.action  = "YSBAE0101.FORE.ASP?totalpage=0"
		<%=formname%>.submit()
	else
		<%=formname%>.action  = "YSBAE0101.FORE.ASP"
		<%=formname%>.submit()
	end if 	
		
end function  

function f()
	<%=formname%>.pid.focus()
end function  

function dataAdd() 
	<%=formname%>.action  = "YSBAE0101.InsertDB.asp?mode=addNew"
	<%=formname%>.submit()
end function  

function godel(index) 
	delPID =  Trim(<%=formname%>.TxtPROGRAM_ID(index).value) 
	if confirm("Delete This Record(Data)?",64) then 
		<%=formname%>.action  = "YSBAE0101.InsertDB.asp?mode=delData&delPid=" & delpid 
		<%=formname%>.submit()
	end if	
end function 

function page_chg()
	<%=formname%>.submit()
end function 


-->
</SCRIPT> 
</head>
<body onkeydown="enterto()"  topmargin=5   onload=f()>
<form NAME ="<%=formname%>" ACTION="<%=action%>" method="<%=method%>" id="<%=formname%>" >
<table width="460" border="0" cellspacing="0" cellpadding="0">
  <tr>
   	<TD ><img border="0" src="../image/icon.gif" align="absmiddle">
   	系統管理(代碼檔維護)</TD>		 
  </tr>
</table> 	
<table border="0"   width=650 CLASS=TXT>
<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>"> 
<hr size=0	style='border: 1px dotted #999999;' align=left width=580>	

	<tr>
		<%if session("RIGHTS")="0" then
			classtype = "INPUTBOX" 
		  else
		    classtype = "READONLY2"
		  end if 	
		%>
		<td >ProgramID</td>
		<td ><input name=pid class="inputBOX" size=8 value="<%=pid%>" onblur=pidchg()></td>		
		<td >LEVEL</td>
		<td ><input name=LEVEL class="<%=classtype%>"  size=7 value="<%=LEVEL%>" <%if  session("RIGHTS")<>"0" then %>readonly <%end if%>></td> 
		<td >UPID</td>
		<td ><input name=UPID class="<%=classtype%>"  size=7 value="<%=UPID%>" <%if  session("RIGHTS")<>"0"  then %>readonly <%end if%> > </td>
		<td >PATH</td>
		<td ><input name=vpath class="<%=classtype%>"  size=30 value="<%=vpath%>" <%if  session("RIGHTS")<>"0" then %>readonly <%end if%> ></td>
	</tr>	
	<tr>
		<td >Pro.Name</td>
		<td colspan=3><input name=pname class="<%=classtype%>" size=35 value="<%=pname%>"  <%if session("RIGHTS")<>"0"  then %>readonly <%end if%> ></td>
		<td colspan=2>Pro.Name(VN)</td>
		<td colspan=2>
			<input name=pnameVN class=INPUTBOX size=25 value="<%=pnameVN%>">
			<input type=button class=button name="BTN2" value="AddNew" onclick=dataadd()>
		</td>
	</tr> 
</table>
<HR SIZE=0 WIDTH=750 ALIGN=LEFT>
<table  width=750  border="0" class="TXT" ID="Table4">
	
	<tr BGCOLOR=#B4C5DA>		
		
		<td align=center>PRO_id</td>
		<td align=center>PRO_NAME</td>
		<td align=center>PRO_NAME(VN)</td>	
		<td align=center>UPID</td>
		<td align=center>LEVEL</td>
		<td align=center>VIRTUAL_PATH</td>
		<td HEIGHT=20 align=center>DEL</td>			
	</tr>
	<%
	for CurrentRow = 1 to PageRec	
		'RESPONSE.WRITE PageRec
		if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then 
      	%> 
		<tr>
			
			<td><input type="text" class="READONLY2" name="TxtPROGRAM_ID"  size="10" maxlength=20 value="<%=(Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1))))%>"  readonly ></td>
			<td><input type="text" class="READONLY2" name="TxtPROGRAM_NAME" size="20" maxlength=50 value="<%=server.HTMLEncode(Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2))))%>" readonly></td>
			<td><input type="text" class="READONLY2" name="TxtPROGRAM_NAME_VN" size="30" maxlength=100 value="<%= (Ucase(trim(tmpRec(CurrentPage, CurrentRow, 6))))%>" readonly></td>			
			<td><input type="text" class="READONLY2" name="TxtLAYER_UP" size="5" maxlength=20 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))%>" readonly></td>
			<td><input type="text" class="READONLY2" name="TxtLAYER" size="5" maxlength=10 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 4)))%>" readonly></td>
			<td><input type="text" class="READONLY2" name="TxtVIRTUAL_PATH" size="25" maxlength=100 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>" readonly></td>
			<td>
				<%if session("RIGHTS")="0" then %>
					<input type=button name=op value="DEL" class=button onclick="godel(<%=currentRow-1%>)" >
				<%else%>	
					<input type=hidden  name=op   >
				<%end if%>
			</td>
			
		</tr>
		<%else%>				
			<input type=hidden name=op>
			<input type=hidden name=TxtPROGRAM_ID>
			<input type=hidden name=TxtPROGRAM_NAME>
			<input type=hidden name=TxtPROGRAM_NAME_VN>
			<input type=hidden name=TxtLAYER_UP>
			<input type=hidden name=TxtLAYER>
			<input type=hidden name=TxtVIRTUAL_PATH>
		<%end if%>
	<%next%>
	
</table>
<input type=hidden name=op>
<input type=hidden name=TxtPROGRAM_ID>
<input type=hidden name=TxtPROGRAM_NAME>
<input type=hidden name=TxtPROGRAM_NAME_VN>
<input type=hidden name=TxtLAYER_UP>
<input type=hidden name=TxtLAYER>
<input type=hidden name=TxtVIRTUAL_PATH>

<table border="0" class=TXT WIDTH=600>	
	<td  align="center" COLSPAN=2>
	Page:
	<select name=topage class=inputbox onchange=page_chg()>
	<%for k= 1 to totalpage %>
		<option value=<%=k%> <%if cdbl(k) = cdbl(CurrentPage) then %> selected <%end if%> > <%=k%></option>
	<%next%>
	</select>
	/ Total Page:<%=TotalPage%> / Total RecordCount:<%=RecordInDB%>
	</td>
	<tr>
		<td width =100%  align =left >
			<INPUT type="SUBMIT" value="第一頁" name="send" class="button" ID="btn_first">
			<INPUT type="SUBMIT" value="上一頁" name="send" class="button" ID="btn_prev">
			<INPUT type="SUBMIT" value="下一頁" name="send" class="button" ID="btn_next">
			<INPUT type="SUBMIT" value="最末頁" name="send" class="button" ID="btn_last">	
			
		</TD>
		<!--TD WIDTH=40% ALIGN=RIGHT>	
			<INPUT type="hidden" value="<%=iRows-1%>" id="cnt" name="cnt">			
			<INPUT type="button" value="確  定" id="BtnUpdate" name="BtnSure" class="button" onclick="go(2)">
			<INPUT type="reset"  value="取  消" id="BtnRst"  name="BtnRst" class="button">
			
		</td-->
	</tr>
</table>

</form>
</body>
</html>
<%
Sub StoreToSession()
	Dim CurrentRow
	tmpRec = Session("SYSPRO01")
	for CurrentRow = 1 to PageRec
		'tmpRec(CurrentPage, CurrentRow, 0) = request("op")(CurrentRow)	 
		
		tmpRec(CurrentPage, CurrentRow, 1) = request("TxtPROGRAM_ID")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 2) = request("TxtPROGRAM_NAME")(CurrentRow)
		tmpRec(CurrentPage, CurrentRow, 3) = request("TxtLAYER_UP")(CurrentRow)		
		tmpRec(CurrentPage, CurrentRow, 4) = request("TxtLAYER")(CurrentRow)	
		tmpRec(CurrentPage, CurrentRow, 5) = request("TxtVIRTUAL_PATH")(CurrentRow)	
		tmpRec(CurrentPage, CurrentRow, 6) = request("TxtPROGRAM_NAME_VN")(CurrentRow)	
	next 
	Session("SYSPRO01") = tmpRec
End Sub
%>
<script language=vbscript> 
function pidchg()
	if <%=formname%>.pid.value<>"" then 
		<%=formname%>.pid.value = Ucase(trim(<%=formname%>.pid.value))
		pidstr= Ucase(trim(<%=formname%>.pid.value))
		open "YSBAE0101.back.asp?func=pidchg&pid=" & pidstr , "Back" 
		'parent.best.cols="70%, 30%"
	end if	
end function 
 
</script>

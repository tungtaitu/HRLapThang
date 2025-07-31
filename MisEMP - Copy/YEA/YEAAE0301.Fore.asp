<%@Language=VBScript Codepage=65001 %>
<!--#include file="../GetSQLServerConnection.fun"--> 
<!--#include file="../include/sideinfo.inc"-->
<%Response.Buffer =True%>  
<%
response.Buffer = true
session.CodePage = "65001"
response.Charset = "utf-8" 

SELF="YEAAE0301"

const action	="YEAAE0301.FORE.ASP"
const formname	="FRM"
const method	="POST"

pid = request("pid")
upid = request("upid")
level = request("level")
vpath = request("vpath")
pname = request("pname")
pnameVN = request("pnameVN") 

Set conn = GetSQLServerConnection()
Set rs = Server.CreateObject("ADODB.Recordset")     

gTotalPage = 50
PageRec = 10   'number of records per page
TableRec = 10    'number of fields per record
Query = TRIM(request("schx"))

IF TRIM(QUERY)="" THEN  	
	if request("TotalPage") = "" or request("TotalPage") = "0" then 
		CurrentPage = 1	  	
	  	Source = "select * from SYSPROGRAM WHERE PROGRAM_ID='"& Query &"' order by PROGRAM_ID"	 
		
	  	rs.Open Source, conn, 3, 3		
	  	IF NOT RS.EOF THEN 			
			RecordInDB = rs.RecordCount 
			pagerec = RecordInDB 
			rs.PageSize = PageRec 
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
	else    
	  	TotalPage = cint(request("TotalPage"))
	  	'response.write "TotalPage=" & TotalPage  &"<BR>"
	  	topage = cint(request("topage"))
		if topage="" then topage=0
		'response.write "topage=" & topage  &"<BR>"
		gTotalPage = request("gTotalPage")
	  	StoreToSession()
	  	CurrentPage = cint(request("CurrentPage"))
	  	RecordInDB = request("RecordInDB")
	  	pagerec = request("pagerec")
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
			RecordInDB = rs.RecordCount 
			pagerec=RecordInDB 
			rs.PageSize = PageRec 			
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
<link REL="stylesheet" HREF="../include/base.css" TYPE="text/css">  
<link REL="stylesheet" HREF="../include/SelectStyle2.css" TYPE="text/css">
<link REL="stylesheet" HREF="../include/Style2.css" TYPE="text/css"> 
<link REL="stylesheet" HREF="../include/Style.css" TYPE="text/css">
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
	if trim(<%=formname%>.schx.value)="" then 
		<%=formname%>.action  = "<%=SELF%>.FORE.ASP?totalpage=0"
		<%=formname%>.submit()
	else
		<%=formname%>.action  = "<%=SELF%>.FORE.ASP"
		<%=formname%>.submit()
	end if 	
		
end function  

function f()
	<%=formname%>.pid.focus()
end function  

function dataAdd() 
	<%=formname%>.action  = "<%=SELF%>.InsertDB.asp?mode=addNew"
	<%=formname%>.submit()
end function  

function godel(index) 
	delPID =  Trim(<%=formname%>.TxtPROGRAM_ID(index).value) 
	if confirm("Delete This Record(Data)?",64) then 
		<%=formname%>.action  = "<%=SELF%>.InsertDB.asp?mode=delData&delPid=" & delpid 
		<%=formname%>.submit()
	end if	
end function 

function page_chg()
	<%=formname%>.submit()
end function
 
function go()
	<%=formname%>.action = "<%=SELF%>.upd.asp"
	<%=formName%>.submit()
end function  


-->
</SCRIPT>  
</head>
<body onkeydown="enterto()" onload=f()>
<form NAME ="<%=formname%>" ACTION="<%=action%>" method="<%=method%>" id="<%=formname%>" >

<INPUT TYPE=hidden NAME=TotalPage VALUE="<%=TotalPage%>">
<INPUT TYPE=hidden NAME=gTotalPage VALUE="<%=gTotalPage%>">
<INPUT TYPE=hidden NAME=CurrentPage VALUE="<%=CurrentPage%>">
<INPUT TYPE=hidden NAME=PageRec VALUE="<%=PageRec%>">
<INPUT TYPE=hidden NAME=RecordInDB VALUE="<%=RecordInDB%>"> 
<table border=0 style="height:70px"><tr><td>&nbsp;</td></tr></table>

	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >					
		<tr>
			<td align="center">
				<table id="myTableForm" width="90%">
					<tr><td colspan=8>&nbsp;</td></tr>
					<tr>
						<%if session("RIGHTS")<>"0" then
							classtype = "inputbox" 
						  else
							classtype = "inputbox"
						  end if 	
						%>
						<td class="frmtable-label">ProgramID</td>
						<td ><input type="text" style="width:100px" name=pid  size=7 value="<%=pid%>" onblur=pidchg()></td>		
						<td class="frmtable-label">LEVEL</td>
						<td ><input type="text" style="width:60px" name=LEVEL  size=4 value="<%=LEVEL%>" <%if session("RIGHTS")<>"0" then %>readonly <%end if%>></td> 
						<td class="frmtable-label">UPID</td>
						<td ><input type="text" style="width:60px" name=UPID  size=5 value="<%=UPID%>" <%if session("RIGHTS")<>"0" then %>readonly <%end if%> > </td>
						<td class="frmtable-label">PATH</td>
						<td ><input type="text" style="width:300px" name=vpath  size=28 value="<%=vpath%>" <%if session("RIGHTS")<>"0" then %>readonly <%end if%> ></td>
					</tr>	
					<tr>
						<td class="frmtable-label">Pro.Name</td>
						<td colspan=3><input type="text" style="width:98%" name=pname  size=18 value="<%=pname%>"  <%if session("RIGHTS")<>"0" then %>readonly <%end if%> ></td>
						<td colspan=2 class="frmtable-label">Pro.Name(VN)</td>
						<td colspan=2>
							<input type="text" style="width:98%" name="pnameVN" id="pnameVN"  size=25 value="<%=pnameVN%>">										
						</td>
					</tr>
					<tr>
						<td align="center" colspan="8" height="40px">
							<input type="button" class="btn btn-sm btn-danger" name="BTN2" id="BTN2" value="Add New" onclick="dataAdd()">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<br>	
	
	<table width="100%" BORDER=0 align=center cellpadding=0 cellspacing=0 >					
		<tr>
			<td align="center">
				<table class="txt" width="98%">
					<tr>
						<td nowrap nowrap>SCH:
							<select name="schx" id="schx" class="form-control-sm" onchange="gosrch()" style="width:200px">
								<option value="" <%if request("schx")="" then %>selected <%end if%>>ALL</option>
								<%
								sql="select * from sysprogram   where  len(program_id)=1 order by program_id "
								set rstmp=conn.execute(sql)
								while not rstmp.eof
								%>
								<option value="<%=rstmp("program_id")%>" <%if request("schx")=rstmp("program_id") then %>selected <%end if%>><%=rstmp("program_id")%>-<%=rstmp("program_name")%></option>
								<%
								rstmp.movenext
								wend
								rstmp.close
								set rstmp=nothing
								%>
							</select>		
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<table id="myTableGrid" width="90%">								
					<tr class="header">
						<td>PRO_id</td>
						<td>PRO_NAME</td>
						<td>PRO_NAME(VN)</td>	
						<td>UPID</td>
						<td>LEVEL</td>
						<td>VIRTUAL_PATH</td>
						<td>DEL</td>	
					</tr>
				<%
				for CurrentRow = 1 to PageRec	
					if trim(tmpRec(CurrentPage, CurrentRow, 1))<>"" then 
				%> 
					<tr>									
						<td><input type="text" style="width:98%"  name="TxtPROGRAM_ID"  size="7" maxlength=20 value="<%=(Ucase(trim(tmpRec(CurrentPage, CurrentRow, 1))))%>"  readonly ></td>
						<td><input type="text" style="width:98%"  name="TxtPROGRAM_NAME" size="16" maxlength=50 value="<%=server.HTMLEncode(Ucase(trim(tmpRec(CurrentPage, CurrentRow, 2))))%>" onchange="datachg(<%=currentrow-1%>)" ></td>
						<td><input type="text" style="width:98%"  name="TxtPROGRAM_NAME_VN" size="30" maxlength=100 value="<%= (Ucase(trim(tmpRec(CurrentPage, CurrentRow, 6))))%>" onchange="datachg(<%=currentrow-1%>)"></td>			
						<td><input type="text" style="width:98%"  name="TxtLAYER_UP" size="5" maxlength=20 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 3)))%>" onchange="datachg(<%=currentrow-1%>)"></td>
						<td><input type="text" style="width:98%"  name="TxtLAYER" size="5" maxlength=10 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 4)))%>" onchange="datachg(<%=currentrow-1%>)"></td>
						<td><input type="text" style="width:98%"  name="TxtVIRTUAL_PATH" size="35" maxlength=100 value="<%=Ucase(trim(tmpRec(CurrentPage, CurrentRow, 5)))%>"  onchange="datachg(<%=currentrow-1%>)"></td>
						<td align="center">
							<%if Session("RIGHTS")="0" then %>
								<input type=button name=op value="DEL" class="btn btn-sm btn-outline-secondary" onclick="godel(<%=currentRow-1%>)" >
							<%else%>	
								<input type=hidden  name=op   >
							<%end if%>	
							<input type=hidden  name=opn   >
						</td>									
					</tr>
					<%else%>	
						<input type=hidden name=op>
						<input type=hidden  name=opn   >
						<input type=hidden name=TxtPROGRAM_ID>
						<input type=hidden name=TxtPROGRAM_NAME>
						<input type=hidden name=TxtPROGRAM_NAME_VN>
						<input type=hidden name=TxtLAYER_UP>
						<input type=hidden name=TxtLAYER>
						<input type=hidden name=TxtVIRTUAL_PATH>
					<%end if%>
				<%next%>								
				</table>							
			</td>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<td align="center">
				<input type=hidden name=op>
				<input type=hidden  name=opn   >
				<input type=hidden name=TxtPROGRAM_ID>
				<input type=hidden name=TxtPROGRAM_NAME>
				<input type=hidden name=TxtPROGRAM_NAME_VN>
				<input type=hidden name=TxtLAYER_UP>
				<input type=hidden name=TxtLAYER>
				<input type=hidden name=TxtVIRTUAL_PATH>
				<table border="0" class="txt" width="98%">	
					<TR>
						<td>Page:
							<select name=topage class="form-control-sm mr-sm-2" onchange=page_chg() style="width:50px">
							<%for k= 1 to totalpage %>
								<option value=<%=k%> <%if cdbl(k) = cdbl(CurrentPage) then %> selected <%end if%> > <%=k%></option>
							<%next%>
							</select>
							/ Total Page:<%=TotalPage%> / Total RecordCount:<%=RecordInDB%>	&nbsp;&nbsp;&nbsp;
							<INPUT type="SUBMIT" value="第一頁" name="send" class="btn btn-sm btn-outline-secondary" ID="btn_first">
							<INPUT type="SUBMIT" value="上一頁" name="send" class="btn btn-sm btn-outline-secondary" ID="btn_prev">
							<INPUT type="SUBMIT" value="下一頁" name="send" class="btn btn-sm btn-outline-secondary" ID="btn_next">
							<INPUT type="SUBMIT" value="最末頁" name="send" class="btn btn-sm btn-outline-secondary" ID="btn_last">
							<INPUT type="hidden" value="<%=iRows-1%>" id="cnt" name="cnt">
						</td>
					</TR>								
					<tr>
						<td align="center" height="60px">										
							<INPUT type="button" value="確  定" id="BtnUpdate" name="BtnSure" class="btn btn-sm btn-danger" onclick="go()">
							<INPUT type="reset"  value="取  消" id="BtnRst"  name="BtnRst" class="btn btn-sm btn-outline-secondary">
						</td>
					</tr>
				</table>
			</td>
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
		open "<%=SELF%>.back.asp?func=pidchg&pid=" & pidstr , "Back" 
		'parent.best.cols="70%, 30%"
	end if	
end function  


function datachg(index)
	<%=formName%>.opn(index).value="UPD"	
	s1=<%=formname%>.TxtPROGRAM_ID(index).value
	s2=<%=formname%>.TxtPROGRAM_NAME(index).value
	s3=<%=formname%>.TxtPROGRAM_NAME_VN(index).value
	s4=<%=formname%>.TxtLAYER_UP(index).value
	s5=<%=formname%>.TxtVIRTUAL_PATH(index).value
	
	open "<%=SELF%>.back.asp?CurrentPage=" & <%=CurrentPage%> & _ 
		 "&program_id=" & S1 & _ 
		 "&program_name=" & S2 & _ 
		 "&PROGRAM_NAME_VN=" & S3 & _ 
		 "&layer_up=" & S4 & _ 
		 "&VIRTUAL_PATH=" & S5 & _ 
		 "&index=" & index & "&func=datachg", "Back"
	
	'parent.best.cols="70%,30%"
	
	
	  
	
end function 
 
</script>

<%@Language=VBScript Codepage=65001%>
<meta HTTP-EQUIV="Content-Type" content="text/html; charset=utf-8">
<!---------  #include file="../GetSQLServerConnection.fun"  -------->

<%
' Set fso = Server.CreateObject("Scripting.FileSystemObject")
' Set fn = fso.GetFolder("uploadfile")
' Set fc = fn.Files
  ' For Each fn in fc 
     ' SourceFile = Server.MapPath(fn.Name) 
     ' TargetFile = Server.MapPath("ytb/upload/XXX.xls")
     ' On Error Resume Next
	' fso.MoveFile SourceFile, TargetFile 
  ' Next 
	fliename="2008_la.xls"

Set ConnXls = Server.CreateObject("ADODB.Connection")
    Driver = "Driver={Microsoft Excel Driver (*.xls)};"
    DBPath = "DBQ=" & Server.MapPath("khb/"& fliename &")"
    ConnXls.Open Driver & DBPath

Set ConnSQL = GetSQLServerConnection() 	
	
response.write connsql
k = 0 
Set rs=ConnXls.execute("Select * from [sheet1$]") 'this [SHEET1$] IS EXCEL WORK SHEET,YOUR CAN CTRL YOUR import SHEET
  Do Until rs.eof		
		k = k + 1  
		groupid = rs.Fields(2) 
		country = rs.Fields(3) 
		empid = rs.Fields(4) 
		indat = rs.Fields(6) 		
		nz = rs.Fields(7) 
		fensu = rs.Fields(8) 
		kj = rs.Fields(9) 
		if K>1 then 
			strsql1="insert into EmpNZKH([years], [whsno], [country],[empid], [indat], [groupid], [nz], [fensu], [kj], [mdtm], [muser]) values ( "&_
					"'"&nam&"','"&whsno&"', '" &country& "','" &empid & "','" &indat & "','" &groupid& "','" &nz& "','" &fensu& "','" &kj& "' ,"&_
					"getdate(),'"& session("netuser") &"' )"
	    'ConnSQL.execute strsql1
			rs.moveNext
			response.write strsql1&"<BR>"
		end if 
	Loop 
 
	
	response.end  	
 
%>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 
<!-- #include file = "../GetSQLServerConnection.fun" --> 
<HEAD>
<META http-equiv="Content-Type" content="text/html; charset=utf-8">
</HEAD>
<%
Set conn = GetSQLServerConnection()	 
Response.Charset="utf-8"
'如果發生錯誤，先跳過
'On Error Resume Next  

'filename=Request.QueryString("H1") 
'response.write "XXXX"
'response.end 
 
%>
<HTML>
<BODY BGCOLOR="white">
<%   
'  	Variables
'  	*********
  Dim mySmartUpload
  Dim intCount  
	intCount=0 
'  	Object creation
'  	***************
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")	
	'response.end 
	
'	Upload  ******
	'mySmartUpload.MaxFileSize = 50000   
	'myUploadfile.AllowedFilesList = "jpg,gif"
	mySmartUpload.Upload 	   
	
	'response.end 
	
   ''支援中文檔名
	'mySmartUpload.Files.Item(1).SaveAs("/aspSmartUpload/Upload/123.jpg")
	'Set fso = CreateObject("Scripting.FileSystemObject")
	'fso.CopyFile Server.MapPath("/aspSmartUpload/Upload/123.jpg"), savefilename
	'fso.DeleteFile Server.MapPath("/aspSmartUpload/Upload/123.jpg")
	'Set fso=Nothing    

	'  Save the files with their original names in a virtual path of the web server
	'  ****************************************************************************
	' intCount = mySmartUpload.Save("UpdData")
	' sample with a physical path 
	' intCount = mySmartUpload.Save("c:\temp\")  

   For each file In mySmartUpload.Files		
			'  Only if the file exist
			'  **********************
			If not file.IsMissing Then
			'  Save the files with his original names in a virtual path of the web server
			'  ****************************************************************************      				
				' sample with a physical path 
				f_filename= file.Name 
				nam=mySmartUpload.Form("years").values
				whsno=mySmartUpload.Form("whsno").values
				'response.write 	intCount  &"<BR>" 
				
				if f_filename = "FILE1" then  
					savefilename = nam&"_"&whsno &"."& file.FileExt 
					file.SaveAs("khb/" & savefilename ) 
				end if  
				'response.write savefilename
				'response.end 
				'file.SaveAs("pic2/" & file.FileName )  '測試
	      intCount = intCount + 1         
	       
	      '  Display the properties of the current file
	      '  ******************************************
				' Response.Write("Name = " & file.Name & "<BR>")
				' Response.Write("Size = " & file.Size & "<BR>")
				' Response.Write("FileName = " & file.FileName & "<BR>")
				' Response.Write("FileExt = " & file.FileExt & "<BR>")
				
				'Response.Write("FilePathName = " & file.FilePathName & "<BR>")
				'Response.Write("ContentType = " & file.ContentType & "<BR>")
				'Response.Write("ContentDisp = " & file.ContentDisp & "<BR>")
				'Response.Write("TypeMIME = " & file.TypeMIME & "<BR>")
				'Response.Write("SubTypeMIME = " & file.SubTypeMIME & "<BR>")           

				relFileName = left(trim(file.FileName),len(file.FileName)-len(file.FileExt)-1 )          
				'sql="update  hwkhbM  set  fileID='"& file.FileName &"' where khid='"&relFileName&"'  " 
				'conn.execute(Sql)
				'Response.Write sql &"<BR>"    
			End If
	Next

'  Display the number of files which could be uploaded
'  *************************************************** 
   	'Response.Write("<BR>" & mySmartUpload.Files.Count & " files could be uploaded.<BR>")        	
'  Display the number of files uploaded
'  ************************************  
'response.write savefilename 
'response.end  
if err.number=0 then 
	Response.Write(intCount & " file(s) uploaded.") &"<BR>"   
	'上傳成功後寫入資料庫 
Set ConnXls = Server.CreateObject("ADODB.Connection")
    Driver = "Driver={Microsoft Excel Driver (*.xls)};"
    DBPath = "DBQ=" & Server.MapPath("khb/")&"\"&savefilename
		response.write "pathe="& Driver & DBPath &"<BR>"
		'response.end 
    ConnXls.Open Driver & DBPath
'response.end
Set ConnSQL = GetSQLServerConnection() 	
ConnSQL.BeginTrans	
'response.write connsql
k = 0  

Set rs=ConnXls.execute("Select * from ["&relFileName&"$]") 'this [SHEET1$] IS EXCEL WORK SHEET,YOUR CAN CTRL YOUR import SHEET
  Do Until rs.eof		
		k = k + 1  
		groupid = rs.Fields(2) 
		country = rs.Fields(3) 
		empid = rs.Fields(4) 
		indat = rs.Fields(6) 		
		nz = rs.Fields(7) 
		fensu = trim(rs.Fields(8)) 
		kj = trim(rs.Fields(9))
		'response.write K & empid &"<BR>"
		if K>1 then  
			sqlx="select * from EmpNZKH where years='"&nam&"' and empid='"& empid &"' "
			Set rsx = Server.CreateObject("ADODB.Recordset") 
			rsx.open sqlx, ConnSQL, 3, 3 
			if rsx.eof then 
				strsql1="insert into EmpNZKH([years], [whsno], [country],[empid],   [fensu], [kj], [mdtm], [muser]) values ( "&_
						"'"&nam&"','"&whsno&"', '" &country& "','" &empid & "', '" &fensu& "','" &kj& "' ,"&_
						"getdate(),'"& session("netuser") &"' )"
				ConnSQL.execute(strsql1)			 
				'response.write strsql1&"<BR>"
			else
				sql1="delete  EmpNZKH where years='"&nam&"' and empid='"& empid &"' "
				ConnSQL.execute(sql1)
				strsql1="insert into EmpNZKH([years], [whsno], [country],[empid],  [fensu], [kj], [mdtm], [muser]) values ( "&_
						"'"&nam&"','"&whsno&"', '" &country& "','" &empid & "', '" &fensu& "','" &kj& "' ,"&_
						"getdate(),'"& session("netuser") &"' )"
				ConnSQL.execute(strsql1) 
				'response.write strsql1&"<BR>"				
			end if 			
			set rsx=nothing 
		end if 
		rs.moveNext
	Loop  
	
	'response.end 
	if err.number = 0 then
		ConnSQL.CommitTrans
		response.write "資料轉入成功!!success(OK)!!" 
		response.redirect "yece1303.fore.asp?flag=S&whsno="&whsno &"&years="& nam 
	else	
		ConnSQL.RollbackTrans 
		response.write "資料轉入失敗,請重新上傳!!Fail(error)!!"&err.description 
	end if 	
	response.end  	


	
	
ELSE
		response.write "檔案上傳錯誤!!<hr>"&"<BR>"
		'response.write err.number &":"& err.description &"<BR>"
		'Response.Write("max Size = " & mySmartUpload.MaxFileSize & "<BR>")
		'Response.Write("File Size = " & file.Size & "<BR>")
		'response.write  "xx=" & mySmartUpload.Files.TotalBytes 
		'response.write "檔案大小不可超過30K"&"<BR>"
		'response.write "檔案只可接受副檔名為 jpg 的檔案 "&"<BR><BR>"
		response.write "<center><input name=sbtn type=button value='關閉視窗close' onclick='window.close()' ></center>"&"<BR>"
		response.end 
end if		
%>
</BODY>
</HTML>
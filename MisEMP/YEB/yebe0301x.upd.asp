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
	'mySmartUpload.MaxFileSize = 30000   
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
				response.write f_filename 
				
				nam="MST"
				whsno=mySmartUpload.Form("whsno").values
				response.write 	whsno  &"<BR>" 
				'response.end 
				'if f_filename = "FILE1" then  
					savefilename = "MST_"&whsno &"."& file.FileExt 
					file.SaveAs("MST/" & savefilename ) 
				'end if  
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

if err.number=0 then 
	Response.Write(intCount & " file(s) uploaded.") &"<BR>"   
response.write savefilename 
'response.end  

	'上傳成功後寫入資料庫 
Set ConnXls = Server.CreateObject("ADODB.Connection")
    Driver = "Driver={Microsoft Excel Driver (*.xls)};"
    DBPath = "DBQ=" & Server.MapPath("MST/")&"\"&savefilename
		response.write "pathe="& Driver & DBPath &"<BR>"
		'response.end 
    ConnXls.Open Driver & DBPath
'response.end
Set ConnSQL = GetSQLServerConnection() 	
ConnSQL.BeginTrans	
'response.write connsql

'response.end 
k = 0  

'Set rs=ConnXls.execute("Select * from ["&relFileName&"$]") 'this [SHEET1$] IS EXCEL WORK SHEET,YOUR CAN CTRL YOUR import SHEET
Set rs=ConnXls.execute("Select * from [Sheet1$]") 'this [SHEET1$] IS EXCEL WORK SHEET,YOUR CAN CTRL YOUR import SHEET
  Do Until rs.eof		
		k = k + 1  
		 
		empid = rs.Fields(0) 
		taxcode = rs.Fields(1) 
		 
		response.write K & empid  &taxcode & "<BR>"
		'response.end 
		 if K>=1 then  
			sqlx="update  empfile set taxCode='"& taxcode &"' where  empid='"& Ucase(trim(empid)) &"' "
			'response.write sqlx &"<BR>"
			ConnSQL.execute(sqlx)
		end if 
		rs.moveNext
	Loop  
	
	'response.end 
	if err.number = 0 then
		ConnSQL.CommitTrans 
		ConnSQL.close
		set ConnSQL=nothing
		%>
		<script language="vbscript">
			alert "資料轉入成功!!success(OK)!!"  
			window.close()
		</script>
<%'		response.write "資料轉入成功!!success(OK)!!" 
		'response.redirect "yece1303.fore.asp?flag=S&whsno="&whsno &"&years="& nam 
	else	
		ConnSQL.RollbackTrans  
		ConnSQL.close
		set ConnSQL=nothing
		%>
		<script language="vbscript">
			alert "資料轉入失敗,請重新上傳!!Fail(error)!!"
			window.close()
		</script>		
<%		response.write "資料轉入失敗,請重新上傳!!Fail(error)!!"&err.description 
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
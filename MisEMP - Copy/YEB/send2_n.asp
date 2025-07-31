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

'empid=request("empid")(1)
'response.write empid 
'response.end 

filename=Request.QueryString("H1")
savefilename=Server.MapPath("/aspSmartUpload/Upload")&"\"&filename

%>

<HTML>
<BODY BGCOLOR="white">

<H1>aspSmartUpload : Sample 1</H1>
<HR>

<% 

'  Variables
'  *********
   Dim mySmartUpload
   Dim intCount  
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

	'response.write mySmartUpload 
	'response.end 
'  Upload
'  ******
   mySmartUpload.Upload 
   
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
         file.SaveAs("UpdData/" & file.FileName)
         ' sample with a physical path 
         ' file.SaveAs("c:\temp\" & file.FileName)

      '  Display the properties of the current file
      '  ******************************************
         Response.Write("Name = " & file.Name & "<BR>")
         Response.Write("Size = " & file.Size & "<BR>")
         Response.Write("FileName = " & file.FileName & "<BR>")
         Response.Write("FileExt = " & file.FileExt & "<BR>")
         'Response.Write("FilePathName = " & file.FilePathName & "<BR>")
         'Response.Write("ContentType = " & file.ContentType & "<BR>")
         'Response.Write("ContentDisp = " & file.ContentDisp & "<BR>")
         'Response.Write("TypeMIME = " & file.TypeMIME & "<BR>")
         'Response.Write("SubTypeMIME = " & file.SubTypeMIME & "<BR>")          
         
         relFileName = left(trim(file.FileName),len(file.FileName)-len(file.FileExt)-1 )          
         sql="update  hwkhbM  set  fileID='"& file.FileName &"' where khid='"&relFileName&"'  " 
         conn.execute(Sql)
         Response.Write sql &"<BR>"
         
         intCount = intCount + 1
      End If
   Next

'  Display the number of files which could be uploaded
'  ***************************************************
   Response.Write("<BR>" & mySmartUpload.Files.Count & " files could be uploaded.<BR>")       

'  Display the number of files uploaded
'  ************************************ 
   if err.number=0 then 
			conn.close
			set conn=nothing
   		Response.Write(intCount & " file(s) uploaded.")    			
%>   		
		<script language=vbscript>
		'	alert  "上傳完成!!"
		'	 parent.close()
		</script>
<% ELSE
		response.end 
   end if		
%>
</BODY>
</HTML>
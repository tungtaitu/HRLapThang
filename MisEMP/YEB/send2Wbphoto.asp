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

filename=Request.QueryString("H1")
savefilename=Server.MapPath("pic/")&"\"&filename
%>

<HTML>
<BODY BGCOLOR="white">
<%  

'  	Variables
'  	*********
   	Dim mySmartUpload
   	Dim intCount  
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
		empid=mySmartUpload.Form("empid").values	  
        'savefilename = empid & ".jpg"          
        file.SaveAs("wbphotos/" & file.FileName ) 
        'file.SaveAs("pic/" & savefilename )        
        
       
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

         intCount = intCount + 1     

         sql="select * from wbphotos where wbempid='"& empid &"'  "
         'response.write sql 
         'response.end 
         set rs=conn.execute(Sql)
         if  rs.eof then 
         	sql="insert into wbphotos ( wbempid, filename ) values ('"& empid&"', '"& file.FileName &"' ) "         	
         else
         	sql="update wbphotos set filename='"& file.FileName  &"' where filename='"& file.FileName  &"' and wbempid='"& empid &"' "	
         end if 
         set rs=nothing 
         response.write sql 
         'conn.execute(Sql)
         set rs=nothing 
		
		Nfname = Server.MapPath("wbphotos")&"\"& file.FileName 
		'response.write Nfname &"<BR>" 		
		'response.end 
		
		Set rsx= Server.CreateObject("ADODB.Recordset") 
		sql="select * from wbempfile where wbid='"& empid &"'"		
		response.write sql &"<BR>"
		
		const adCmdText=1
		const adOpenDynamic=2
		const adLockOptimistic=3
		const adOpenKeyset=1
		
		Set mstream = Server.CreateObject("ADODB.Stream")
		mstream.Type = 1
		mstream.Open
		
		mstream.LoadFromFile Nfname  
		
		'response.end 	
		
		rsx.Open SQL,conn,adOpenKeyset,adLockOptimistic,adCmdText
		if not rsx.eof then
			rsx.Fields("filename").Value = file.FileName 
			rsx.Fields("photos").Value = mstream.Read
			rsx.Update
		else
			rsx.addnew  	
			rsx.Fields("filename").Value = file.FileName 	        
			rsx.Fields("photos").Value = mstream.Read
			rsx.Update        
		end if 
		rsx.close  
         
      End If
   Next

'  Display the number of files which could be uploaded
'  ***************************************************
   	
   	Response.Write("<BR>" & mySmartUpload.Files.Count & " files could be uploaded.<BR>")       
	
'  Display the number of files uploaded
'  ************************************ 
 'response.end 

   if err.number=0 then 
			conn.close
			set conn=nothing
   		Response.Write(intCount & " file(s) uploaded.")    			
%>   		
		<script language=vbscript>
			alert  "上傳完成!!"
			 window.close()
		</script>
<% ELSE
		conn.close
		set conn=nothing
		response.write "檔案上傳錯誤!!<hr>"&"<BR>"
		'response.write err.number &":"& err.description &"<BR>"
		'Response.Write("max Size = " & mySmartUpload.MaxFileSize & "<BR>")
		'Response.Write("File Size = " & file.Size & "<BR>")
		'response.write  "xx=" & mySmartUpload.Files.TotalBytes 
		response.write "檔案大小不可超過30K"&"<BR>"
		response.write "檔案只可接受副檔名為 jpg 的檔案 "&"<BR><BR>"
		response.write "<center><input name=sbtn type=button value='關閉視窗close' onclick='window.close()' ></center>"&"<BR>"
		response.end 
   end if		
%>
</BODY>
</HTML>
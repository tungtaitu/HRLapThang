<%@ Language=VBScript %>
<%

Set objFSO = CreateObject("Scripting.FileSystemObject")
' create a Drives collection
Set colDrives = objFSO.Drives
' iterate through the Drives collection
For Each objDrive in colDrives

  Response.Write "DriveLetter: <B>" & objDrive.DriveLetter & "</B>   "
  Response.Write "DriveType: <B>" & objDrive.DriveType
  Select Case objDrive.DriveType
    Case 0: Response.Write " - (Unknown)"
    Case 1: Response.Write " - (Removable)"
    Case 2: Response.Write " - (Fixed)"
    Case 3: Response.Write " - (Network)"
    Case 4: Response.Write " - (CDRom)"
    Case 5: Response.Write " - (RamDisk)"
  End Select
  Response.Write "</B>   "

If objDrive.DriveType = 3 Then
    If objDrive.IsReady Then
      Response.Write "Remote drive with ShareName: <B>" & objDrive.ShareName & "</B>"
    Else
	Response.Write "Remote drive - <B>IsReady</B> property returned<B>False</B><BR>"
    End If
  Else If objDrive.IsReady then 
    Response.Write "FileSystem: <B>" & objDrive.FileSystem & "</B>   "
    Response.Write "SerialNumber: <B>" & objDrive.SerialNumber & "</B><BR>"
	Response.Write "Local drive with VolumeName: <B>" & objDrive.VolumeName & "</B><BR>"
	Response.Write "AvailableSpace: <B>" & FormatNumber(objDrive.AvailableSpace / 1024, 0) & "</B> KB   "
	Response.Write "FreeSpace: <B>" & FormatNumber( objDrive.FreeSpace / 1024, 0) & "</B> KB   "
	Response.Write "TotalSize: <B>" & FormatNumber( objDrive.TotalSize / 1024, 0) & "</B> KB"
  End if  
  Response.Write "<P>"
  End if
Next
 

'filename="F:\temp\ar200602220802.dat" 
'response.write filename  

'	Set fs = CreateObject("Scripting.FileSystemObject")
	
'	Set fileContent = fs.OpenTextFile(filename,1) 
'	set aaa= fs.GetDrive(jbc) 
response.end 

dat1 = trim(request("dat1"))
dat2 = trim(request("dat2"))

D1 = replace(dat1, "/", "" )
D2 = replace(dat2, "/", "" )


Set fs = CreateObject("Scripting.FileSystemObject")
	Set fldr = FS.GetFolder("F:\\JBC")
	set theFiles = fldr.files
	For Each x In theFiles
		filename = x.name 
		 			
 		response.write  filename  &"<BR>"
 		fnamestr = "F:\JBC\" & filename
		response.write  		fnamestr  &"<BR>"
 		'Set fileContent = fs.OpenTextFile(fnamestr,1)
 		
 		'Do while not fileContent.AtEndOfLine 
 		'	data = fileContent.readline
 		'	wk_LineStr = RTrim(data) 
 		'	For j = 1 To len(wk_LineStr)
 		'		'response.write wk_LineStr &"<BR>" 
 		'		 If Mid(wk_LineStr, j, 1) = "," Then
         '           strArray(i) = Mid(wk_LineStr, k + 1, j - k - 1)
          '         Response.Write i & "-" &strArray(i) &"<br>"
           '         i = i + 1
            '        k = j
             '   End If
                'if j= len(wk_LineStr) then
                '  strArray(i) =Mid(wk_LineStr, k + 1, j-k)
                'end if   
 			'next 	
 		
 		'Loop   
		'fileContent.Close  		
		
 	NEXT

	
	
    
%>	 
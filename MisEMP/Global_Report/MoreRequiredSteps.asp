<%
'====================================================================================
' Retrieve the Records and Create the "Page on Demand" Engine Object
'====================================================================================

On Error Resume Next

session("oRpt").ReadRecords

If Err.Number <> 0 Then                                               
  Response.Write "Error Occurred Reading Records: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Abandon
  Response.End
Else
  If IsObject(session("oPageEngine")) Then                              
  	set session("oPageEngine") = nothing
  End If
  set session("oPageEngine") = session("oRpt").PageEngine
End If
%>
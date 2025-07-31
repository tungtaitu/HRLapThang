

<%const rpt_path="http://172.22.166.23/yfyemp/global_report/"%> 
<HTML>
<HEAD>
<TITLE>Crystal Reports ActiveX Viewer</TITLE>
</HEAD>
<BODY BGCOLOR=C6C6C6 topmargin=0 leftmargin=0>   
<OBJECT ID="CRViewer"
	CLASSID="CLSID:2DEF4530-8CE6-41c9-84B6-A54536C90213"
	WIDTH=100% HEIGHT=97%
	CODEBASE="<%=rpt_path%>activexviewer.cab#Version=9,2,0,442" VIEWASTEXT>
<PARAM NAME="EnableRefreshButton" VALUE=0>
<PARAM NAME="EnableGroupTree" VALUE=0>
<PARAM NAME="DisplayGroupTree" VALUE=0>
<PARAM NAME="EnablePrintButton" VALUE=1>
<PARAM NAME="EnableExportButton" VALUE=1>
<PARAM NAME="EnableDrillDown" VALUE=1>
<PARAM NAME="EnableSearchControl" VALUE=1>
<PARAM NAME="EnableAnimationControl" VALUE=1>
<PARAM NAME="EnableZoomControl" VALUE=1> 
</OBJECT>

 
<SCRIPT LANGUAGE="VBScript">
<!--
Sub Window_Onload
	On Error Resume Next
	Dim webBroker
	Set webBroker = CreateObject("WebReportBroker9.WebReportBroker")
	If ScriptEngineMajorVersion < 2 Then
		window.alert "IE 3.02 users need to get the latest version of VBScript or install IE 4.01 SP1 or newer. Users of Windows 95 additionally need DCOM95.  These files are available at Microsoft's web site."
	else
		Dim webSource
		Set webSource = CreateObject("WebReportSource9.WebReportSource")
		webSource.ReportSource = webBroker
		webSource.URL = "<%=rpt_path%>rptserver.asp"
		webSource.PromptOnRefresh = True
		CRViewer.ReportSource = webSource
	End If
	CRViewer.ViewReport
End Sub
-->
</SCRIPT> 

<OBJECT ID="ReportSource"
	CLASSID="CLSID:934CC260-C5AA-43C4-A657-7B70C5B3DAE1"
	HEIGHT=1% WIDTH=1%
    CODEBASE="<%=rpt_path%>activexviewer.cab#Version=9,2,0,442">
</OBJECT>
<OBJECT ID="ViewHelp"
	CLASSID="CLSID:4B5C9C28-3806-47b5-89A9-93063323160F"
	HEIGHT=1% WIDTH=1%
    CODEBASE="<%=rpt_path%>activexviewer.cab#Version=9,2,0,442">
</OBJECT>  
<div>
<!-- This empty div prevents IE from showing a bunch of empty space for the controls above. -->
</div>

</BODY>
</HTML>

<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 


<!DOCTYPE html>
<html lang="en-US">
<head>
<meta http-equiv="content-type" content="text/html; charset=UTF-8">

<title>人事薪資系統</title>
<meta name="viewport" content="width=device-width">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<script src="template/js/jQuery-3.3.1.js"></script>
<script src="template/bootstrap/js/bootstrap.min.js"></script>
<script src="template/js/sidebar-menu.js"></script>
<link rel="stylesheet" type="text/css" href="template/bootstrap/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="template/font-awesome/css/font-awesome.css">
<link rel="stylesheet" type="text/css" href="template/css/mis.css">
<script async type="text/javascript" src="//static.h-bid.com/w3schools.com/20200121/snhb-w3schools.com.min.js"></script>
<!--Rool can su dung=================================================-->
<script>
if (window.addEventListener) {              
    window.addEventListener("resize", browserResize);
} else if (window.attachEvent) {                 
    window.attachEvent("onresize", browserResize);
}
var xbeforeResize = window.innerWidth;

function browserResize() {
    var afterResize = window.innerWidth;
    if ((xbeforeResize < (970) && afterResize >= (970)) || (xbeforeResize >= (970) && afterResize < (970)) ||
        (xbeforeResize < (728) && afterResize >= (728)) || (xbeforeResize >= (728) && afterResize < (728)) ||
        (xbeforeResize < (468) && afterResize >= (468)) ||(xbeforeResize >= (468) && afterResize < (468))) {
        xbeforeResize = afterResize;
        
        snhb.queue.push(function(){  snhb.startAuction(["try_it_leaderboard"]); });
         
    }	
    fixDragBtn();   
}

$(window).resize(function () {      
      $('#textareacontainer').css("width", "20%");
	  $('#dragbar').css("left", "20%");
	  $('#iframecontainer').css("width", "80%");
  });
  
</script>
<!--=================================================-->
<style>
* {
  -webkit-box-sizing: border-box;
  -moz-box-sizing: border-box;
  box-sizing: border-box;
}
body {
  color:#000000;
  margin:0px;
  font-size:100%;
}
.trytopnav {
  height:40px;
  overflow:hidden;
  min-width:380px;
  position:absolute;
  width:100%;
  top:99px;
  background-color:#f1f1f1;
}
.trytopnav a {
  color:#999999;
}
.w3-bar .w3-bar-item:hover {
  color:#757575 !important;
}
a.w3schoolslink {
  padding:0 !important;
  display:inline !important;
}
a.w3schoolslink:hover,a.w3schoolslink:active {
  text-decoration:underline !important;
  background-color:transparent !important;
}
#dragbar{
  position:absolute;
  cursor: col-resize;
  z-index:3;
  padding:0px;
}
#shield {
  display:none;
  top:0;
  left:0;
  width:100%;
  position:absolute;
  height:100%;
  z-index:4;
}
#framesize span {
  font-family:Consolas, monospace;
}
/*Steven 2020/04/25 sua top*/
#container {
  background-color:#f1f1f1;
  width:100%;
  overflow:auto;
  position:absolute;
  top:50px;
  bottom:0;
  height:auto;
}
/*Steven 2020/04/25 sua width*/
#textareacontainer {
  float:left;
  height:100%;
  width:20%;
  white-space: nowrap;
  
  
}
#iframecontainer {
  float:left;
  height:100%;
  width:80%;
  
}
#textarea, #iframe {
  height:100%;
  width:100%;
  padding-bottom:10px;
  padding-top:1px;
}
#textarea {
  padding-left:10px;
  padding-right:5px;  
}
#iframe {
  padding-left:5px;
  padding-right:10px;  
}
#textareawrapper {
  width:100%;
  height:100%;
  background-color: #ffffff;
  position:relative;
  box-shadow:0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
}
#iframewrapper {
  width:100%;
  height:100%;
  -webkit-overflow-scrolling: touch;
  background-color: #ffffff;
  box-shadow:0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.24);
}
#textareaCode {
  background-color: #ffffff;
  font-family: consolas,"courier new",monospace;
  font-size:15px;
  height:100%;
  width:100%;
  padding:0px;
  resize: none;
  border:none;
  line-height:normal; 
  
   
}
.CodeMirror.cm-s-default {
  line-height:normal;
  padding: 4px;
  height:100%;
  width:100%;
}
#iframeResult, #iframeSource {
  background-color: #ffffff;
  height:100%;
  width:100%;  
}

#textareacontainer.horizontal,#iframecontainer.horizontal{
  height:100%;
  float:none;
  width:100%;
}
#textarea.horizontal{
  height:100%;
  padding-left:10px;
  padding-right:10px;
}
#iframe.horizontal{
  height:100%;
  padding-right:10px;
  padding-bottom:10px;
  padding-left:10px;  
}
#container.horizontal{
  min-height:200px;
  margin-left:0;
}
#tryitLeaderboard {
  overflow:hidden;
  text-align:center;
  margin-top:5px;
  height:90px;
}
  
#iframewrapper {
	
}

body.darktheme {
  background-color:rgb(40, 44, 52);
}
body.darktheme #tryitLeaderboard{
  background-color:rgb(40, 44, 52);
}
body.darktheme .trytopnav{
  background-color:#616161;
  color:#dddddd;
}
body.darktheme #container {
  background-color:#616161;
}
body.darktheme .trytopnav a {
  color:#dddddd;
}
body.darktheme ::-webkit-scrollbar {width:12px;height:3px;}
body.darktheme ::-webkit-scrollbar-track-piece {background-color:#000;}
body.darktheme ::-webkit-scrollbar-thumb {height:50px;background-color: #616161;}
body.darktheme ::-webkit-scrollbar-thumb:hover {background-color: #aaaaaa;}
</style>
<!--[if lt IE 8]>
<style>
#textareacontainer, #iframecontainer {width:48%;}
#container {height:500px;}
#textarea, #iframe {width:90%;height:450px;}
#textareaCode, #iframeResult {height:450px;}
.stack {display:none;}
</style>
<![endif]-->
</head>
<body>
<div id="menuOverlay" class="w3-overlay w3-transparent" style="z-index:4;padding-top:10px;padding-bottom:11px;background-color:#f9f9f9 !important">
	<table border=0 width="100%">
		<tr>
			<td>
				<span class="navbar-brand" style="padding:0px 10px 0px 20px;">			
					<img src="template/img/logohoaduong4.png" height="30" class="d-inline-block align-top"><span class="text-danger"> <b>和唐</b> HÒA ĐƯỜNG</span>
				</span>  
			</td>
			<td align="right">登入者 : <%=Session("netuser")%>(<%=REMOTE_IP%>)</td>
			<td style="width:250px" align="center">
				<a href="main.asp" target="main" class="btn btn-warning btn-sm text-white rounded-pill"><span class="fa fa-home" style="width:60px">訊息頁</span></a>
				<a href="default.asp?logout=Y" target="_top" class="btn btn-warning btn-sm text-white rounded-pill"><span class="fa fa-sign-out" style="width:60px">登出</span></a>
			</td>
		</tr>
	</table>
</div>
<div class="sidebar-toggle" id="btShow">
	<button type="button" class="btn btn-danger btn-sm btn-sidebar-toggle" id="btnopenside" onClick="openSidebar()" style="background-color:#9F0000; display:none;">
		<span class="fa fa-arrow-right" style="width:14px;"></span>
	</button>
	<button type="button" class="btn btn-danger btn-sm btn-sidebar-toggle" id="btncloseside" onClick="closesidebar()" style="background-color:#9F0000;">
		<span class="fa fa-arrow-left"></span>
	</button>
</div>
<input id="hidcontainerleft" type="hidden" />
<input id="hidcontainerright" type="hidden" />

<div id="shield"></div>
<a href="javascript:void(0)" id="dragbar"></a>
<div id="container">	
  <div id="textareacontainer" style="z-index:1;">
    <div id="textarea">		
      <div id="textareawrapper">		
		<iframe name="contents" id="textareaCode" target="main"  src="function.asp?program_id=A" frameborder="0"></iframe>		
      </div>
    </div>
  </div>
  <div id="iframecontainer" style="z-index:2;">
    <div id="iframe">
      <div id="iframewrapper" >
		<iframe name="main" id="iframeContent" style="width:100%;height:100%;" src="main.asp" frameborder="0"></iframe>	  
	  </div>
    </div>
  </div>
</div>

<script>

var dragging = false;
var stack;
function fixDragBtn() {
  var leftpadding, dragleft, containertop, buttonwidth
  var containertop = Number(w3_getStyleValue(document.getElementById("container"), "top").replace("px", ""));
  
	document.getElementById("dragbar").style.width = "5px";    
    textareasize = Number(w3_getStyleValue(document.getElementById("textareawrapper"), "width").replace("px", ""));
    leftpadding = Number(w3_getStyleValue(document.getElementById("textarea"), "padding-left").replace("px", ""));
    buttonwidth = Number(w3_getStyleValue(document.getElementById("dragbar"), "width").replace("px", ""));
    textareaheight = w3_getStyleValue(document.getElementById("textareawrapper"), "height");
    dragleft = textareasize + leftpadding + (leftpadding / 2) - (buttonwidth / 2);
    document.getElementById("dragbar").style.top = containertop + "px";
    document.getElementById("dragbar").style.left = dragleft + "px";
    document.getElementById("dragbar").style.height = textareaheight;
    document.getElementById("dragbar").style.cursor = "col-resize";
}

function dragstart(e) {
  e.preventDefault();
  dragging = true;
  var main = document.getElementById("iframecontainer");
}

function dragmove(e) {
  if (dragging) 
  {
    document.getElementById("shield").style.display = "block";        
    if (stack != " horizontal") {
      var percentage = (e.pageX / window.innerWidth) * 100;
      if (percentage > 5 && percentage < 98) {
        var mainPercentage = 100-percentage;
        document.getElementById("textareacontainer").style.width = percentage + "%";
        document.getElementById("iframecontainer").style.width = mainPercentage + "%";
        fixDragBtn();
      }
    } else {
      var containertop = Number(w3_getStyleValue(document.getElementById("container"), "top").replace("px", ""));
      var percentage = ((e.pageY - containertop + 20) / (window.innerHeight - containertop + 20)) * 100;
      if (percentage > 5 && percentage < 98) {
        var mainPercentage = 100-percentage;
        document.getElementById("textareacontainer").style.height = percentage + "%";
        document.getElementById("iframecontainer").style.height = mainPercentage + "%";
        fixDragBtn();
      }
    }   
  }
}

function dragend() {
  document.getElementById("shield").style.display = "none";
  dragging = false;
  var vend = navigator.vendor;
  if (window.editor && vend.indexOf("Apple") == -1) {
      window.editor.refresh();
  }
}
if (window.addEventListener) {              
  document.getElementById("dragbar").addEventListener("mousedown", function(e) {dragstart(e);});
  document.getElementById("dragbar").addEventListener("touchstart", function(e) {dragstart(e);});
  window.addEventListener("mousemove", function(e) {dragmove(e);});
  window.addEventListener("touchmove", function(e) {dragmove(e);});
  window.addEventListener("mouseup", dragend);
  window.addEventListener("touchend", dragend);
  window.addEventListener("load", fixDragBtn);
}

function w3_getStyleValue(elmnt,style) {
    if (window.getComputedStyle) {
        return window.getComputedStyle(elmnt,null).getPropertyValue(style);
    } else {
        return elmnt.currentStyle[style];
    }
}

function openSidebar(){
	var hidcontainerleft = $("#hidcontainerleft").val();
	var hidcontainerright = $("#hidcontainerright").val();	
	
	document.getElementById("dragbar").style.width = "5px";	
	document.getElementById("textareacontainer").style.width = hidcontainerleft+"px";
	document.getElementById("iframecontainer").style.width = hidcontainerright+"px";
	
	
	document.getElementById("btnopenside").style.display = "none";
	document.getElementById("btncloseside").style.display = "block";
	
 	
}
function closesidebar(){
	$("#hidcontainerleft").val($("#textareacontainer").width());
	$("#hidcontainerright").val($("#iframecontainer").width());
	
	document.getElementById("dragbar").style.width = "0";
	
	document.getElementById("textareacontainer").style.width = "0";		
	document.getElementById("iframecontainer").style.width = "100%";	
	document.getElementById("iframecontainer").style.marginLeft= "0";	
	
	document.getElementById("btnopenside").style.display = "block";
	document.getElementById("btncloseside").style.display = "none";
	
}
</script>

</body>
</html> 
 

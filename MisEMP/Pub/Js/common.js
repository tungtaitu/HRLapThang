// JScript 檔


function BtnMouseOver(obj,strStyleName){
    obj.className = strStyleName;
}

function BtnMouseOut(obj,strStyleName){
    obj.className = strStyleName;
}
function doSection(secNum) 
{
	if (secNum.className=="off") {
		secNum.className="on";
	} else {
		secNum.className="off";
	}
}
var SysAps = new Array(); // Phi Add 29_8_2011
function doViewAjax2(oView,pageLayer,ajaxFunction)// Phi Add 29_8_2011
{
    
	doSection(oView);
	if (!IsGened(oView.id))
	{
		//oView.children[0].style.display = "block";
		oView.children[0].innerHTML = "<img src='../../Images/loading.gif'><font style='font-size:8pt;color:purple;font-family:Arial'>Reading Data......</font>";
		eval(ajaxFunction);
	}
}

function doViewAjax2a(oView,pageLayer,ajaxFuntion1)// Phi Add 9_9_2011
{
    
	doSection(oView);
	if (!IsGened(oView.id))
	{
		// oView.children[0].style.display = "block";// Edit follow Trung Loi By Phi - 20111006 - Use Fire fox
		oView.children[0].innerHTML = "<img src='../../Images/loading.gif'><font style='font-size:8pt;color:purple;font-family:Arial'>Reading Data......</font>";
		eval(ajaxFuntion1);
	}
}

function doCallBack(res)// Phi Add 29_8_2011
{

	if(res != null && typeof(res) == 'object'){
		var PccTable = res.value;
		var tdid = document.all[PccTable.ViewID];
		tdid.children[0].style.display = "";		// Edit follow Trung Loi By Phi - 20111006 - Use Fire fox
		tdid.children[0].innerHTML = PccTable.TableHtml;
		SysAps[SysAps.length] = PccTable.ViewID; 
	}
}


function IsGened(sViewID)// Phi Add 29_8_2011
{
	for(i = 0; i < SysAps.length ; i++)
	{
		if (SysAps[i] == sViewID)
		{
			return true;
		}
	}
	return false;
}



function MouseOver_Click(objThis)
{
	var orgSrc = objThis.children[0].src;
	var i = orgSrc.lastIndexOf(".gif");
	var strBegin = orgSrc.substr(0,i);
	var strEnd = orgSrc.substr(i);
	
	objThis.children[0].src = strBegin + "_over" + strEnd;
	
}

function MouseOut_Click(objThis)
{
	var orgSrc = objThis.children[0].src;
	var i = orgSrc.lastIndexOf("_over.gif");
	var strBegin = orgSrc.substr(0,i);
	var strEnd = orgSrc.substr(i + 5);
	
	objThis.children[0].src = strBegin + strEnd;
}


function doSectionMenu(secNum,objThis) {
	if (secNum.className=="off") {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/N_Open.gif";
		secNum.className="on";
	} else {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/N_Close.gif";
		secNum.className="off";
	}
}

function doSectionMenuManage(secNum,objThis) {
	if (secNum.className=="off") {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/Y_Open.gif";
		secNum.className="on";
	} else {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/Y_Close.gif";
		secNum.className="off";
	}
}

function doSectionOtherMenu(secNum,objThis,name) {
	if (secNum.className=="off") {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/" + name +"_Open.gif";
		secNum.className="on";
	} else {
		objThis.children(0).src = objThis.children(0).src.substr(0,objThis.children(0).src.lastIndexOf("/")) + "/" + name + "_Close.gif";
		secNum.className="off";
	}
}

function SendMenuAuthToObj(objThis,srcObj)
{
	//alert(objThis.parentElement.parentElement.children(1).children(0).value);
	srcObj.value = objThis.parentElement.parentElement.children(1).children(0).value; 
	alert(srcObj.value);
	
}

function MM_swapImgRestore() { //v3.0
	var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
	var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
	var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
	if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
	var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
	d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
	if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
	var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
	if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}  

function SecondLayer(XMLStr, APNode)
{
	var xmldom  = new ActiveXObject("Microsoft.xmldom");
	xmldom.validateOnParse = true;
	xmldom.async = false;
	xmldom.loadXML(XMLStr);
	var Authorize = xmldom.selectNodes("/PccMsg/Authorize");
	var ap_node = new Array();

	for(var i=0; i < Authorize.length; i++)
	{
		ap_node[i] = Authorize[i].selectSingleNode(APNode).text;
	}
	return ap_node;
}

function GetApLink(XMLStr, ApNode,ApNode1,vAPName)
{
				
	var xmldom  = new ActiveXObject("Microsoft.xmldom");
	xmldom.validateOnParse = true;
	xmldom.async = false;
	xmldom.loadXML(XMLStr);
	var Authorize = xmldom.selectNodes("/PccMsg/Authorize");
	var apName = "";
	var retApLink = "";

	for(var i=0; i < Authorize.length; i++)
	{
		apName = Authorize[i].selectSingleNode(ApNode).text;
		if (apName == vAPName)
		{
			retApLink = Authorize[i].selectSingleNode(ApNode1).text;
			break;
		}
	}
	return retApLink;
}

function GetLayer(XMLStr, ApNode, MenuNode)
{
	var xmldom  = new ActiveXObject("Microsoft.xmldom");
	xmldom.validateOnParse = true;
	xmldom.async = false;
	xmldom.loadXML(XMLStr);
	var Authorize = xmldom.selectNodes("/PccMsg/Authorize");
	
	var menu_nodeA = new Array();
	var records;
	
	for(var i=0; i < Authorize.length; i++)
	{
		var menu_nodeB = new Array();
		records = Authorize[i].selectNodes(ApNode);
		for(var j=0; j < records.length; j++)
		{
			menu_nodeB[j] = records[j].selectSingleNode(MenuNode).text;						
		}
		menu_nodeA[i] = menu_nodeB;
	}
	return menu_nodeA;
}
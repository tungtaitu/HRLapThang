/*
Name:		MIS 專用前端模塊
Author:		Dragon.Zou(SWG)
Date:		2002/8/28
Last:		2002/12/23
CopyRight:	LS	2001-2002
Warning:	NOT ALLOWED TO MODIFY !
*/

var	SCROLL_X=0,SCROLL_Y=0;
var	SCROLL_L=0,SCROLL_R=0,SCROLL_T=0,SCROLL_B=0;
var	LIBRARY_PATH="";
//Get Library Path
for	(var i=0;i<document.scripts.length;i++)
	{
		var iPos=0;
		var sP=document.scripts[i].src;
		if (sP&&(iPos=sP.toUpperCase().indexOf("MISTOOLS.JS"))>=0)
		{
			LIBRARY_PATH=sP.substr(0,iPos);
			break;
		}
	}

//取消數值為0的鏈接
function removeHrefByZero()
{
	for (var i=0;i<document.links.length;)
	if ((""+document.links[i].innerText)=="0") 
		document.links[i].removeAttribute("href");
	else i++;
}

//取消右擊菜單
function disableContextMenu()
{
	document.oncontextmenu=new Function("return false;");	
}

//調整物件位置及大小
function resizeControl(obj,x,y,w,h,bMaximize)
{
	var iTotalX=0,iTotalY=0,sUnit="";
	var bDialog=typeof(window.dialogWidth)=="object";
	if (bMaximize)
	{
	   window.moveTo( 0, 0 );
	   window.resizeTo( screen.availWidth, screen.availHeight );
	}
	iTotalX=(bDialog?window.dialogWidth : document.body.clientWidth)- SCROLL_R - SCROLL_L;
	iTotalY=(bDialog?window.dialogHeight :document.body.clientHeight)- SCROLL_B - SCROLL_T;
	if (SCROLL_X>0 && SCROLL_X<1) SCROLL_X=parseInt(iTotalX*SCROLL_X);
	if (SCROLL_Y>0 && SCROLL_Y<1) SCROLL_Y=parseInt(iTotalY*SCROLL_Y);
	if (SCROLL_X<0) SCROLL_X=iTotalX;
	if (SCROLL_Y<0) SCROLL_Y=iTotalY;
	if (x<0) x=(x==-1)?0:(SCROLL_X +2);
	if (y<0) y=(y==-1)?0:(SCROLL_Y);
	if (w<0) w=(w==-1)?(iTotalX - x ):(SCROLL_X - x);
	if (h<0) h=(h==-1)?(iTotalY - y ):(SCROLL_Y - y);
	if (!obj.style) return;
	obj.style.position="absolute";
	if (w<=0||h<=0) {w=0;h=0;obj.style.display="none";}
	obj.style.top=y+SCROLL_T;
	obj.style.left=x+SCROLL_L;
	obj.style.width=w;
	obj.style.height=h;
}

//以最大化打開窗體
function openWindow( aURL, aWinName,sOptions)
{
   var wOpen;
	if (typeof(sOptions)=="undefined")
   var sOptions = 'status=no,menubar=no,scrollbars=no,resizable=no,toolbar=no'
   	 +	',width=' + (screen.availWidth - 10).toString()
   	 +  	',height=' + (screen.availHeight - 122).toString()
   	 +	',screenX=0,screenY=0,left=0,top=0';
   wOpen = window.open( '', aWinName, sOptions );
   wOpen.location = aURL;
   wOpen.focus();
   wOpen.moveTo( 0, 0 );
   wOpen.resizeTo( screen.availWidth, screen.availHeight );
   return wOpen;
}


function maximizeWindow(wOpen)
{
	wOpen.moveTo( 0, 0 );
    	wOpen.resizeTo( screen.availWidth, screen.availHeight );
}

function calendar(t) {
	var sPath = "Library/calendar1.htm";
	strFeatures = "dialogWidth=206px;dialogHeight=206px;center=yes;help=no;status=no";
	st = t.value;
	sDate = showModalDialog(sPath,st,strFeatures);
	if (sDate)
		{
			t.value = formatDate(sDate,false);
			document.forms[0].submit();
		}
	return false;
}

function formatDate(sDate,flag) {
	var sScrap = "";
	var dScrap = new Date(sDate);
	if (dScrap == "NaN") return sScrap;
	
	iDay = dScrap.getDate();
	iMon = dScrap.getMonth() + 1;
	iYea = dScrap.getFullYear();
        if (flag==true)
		sScrap=""+iYea+(iMon>9?iMon:("0"+iMon))+(iDay>9?iDay:("0"+iDay));
	else
		sScrap = iYea + "/" + iMon + "/" + iDay ;
	return sScrap;
}

function isYScroll(objDiv)
{
	if(objDiv.scrollHeight<=objDiv.offsetHeight)	return false;
	return true;
}

function fixedColumn(objTable,colNO)
{
	var objDiv=objTable.parentElement;
	if(objDiv.scrollWidth<objDiv.offsetWidth)
		return false;
	var iPadding=objTable.cellPadding*2+parseInt(objTable.style.borderWidth);
	var newTableHead=document.createElement("TABLE");
	newTableHead.mergeAttributes(objTable);
	objDiv.parentElement.insertBefore(newTableHead);
	newTableHead.style.position="absolute";
	newTableHead.style.zIndex=1000;
	newTableHead.style.pixelTop=objDiv.offsetTop;
	newTableHead.style.pixelLeft=objDiv.offsetLeft;
	for(var i=0;i<objTable.rows[0].cells.length;i++)
	{
		//objTable.rows[0].cells[i].style.pixelWidth=objTable.rows[0].cells[i].offsetWidth-iPadding;
		objTable.rows[0].cells[i].style.pixelHeight=objTable.rows[0].cells[i].offsetHeight;
	}
	var newTableWidth;
	for(var i=0;i<objTable.rows.length;i++)
	{
		newRow=newTableHead.insertRow();
		newRow.mergeAttributes(objTable.rows[i]);
		newTableWidth=0;
		for(var j=0;j<colNO;j++)
		{
			var objCell=objTable.rows[i].cells[j].cloneNode(true);
			objCell.rowSpan=1;
			newTableWidth+=objTable.rows[i].cells[j].offsetWidth;
			j+=objTable.rows[i].cells[j].colSpan-1;
			newRow.insertBefore(objCell);
				
		}
		//debugger;
		if(objTable.rows[i].cells[0].rowSpan>1)
			i+=objTable.rows[i].cells[0].rowSpan-1;
	}
	newTableHead.style.pixelWidth=newTableWidth ;
}

function fixedTableRow(objTable,topRow,leftRow)
{	
	var objDiv=objTable.parentElement;
	if(objDiv.scrollHeight<objDiv.offsetHeight)
		return false;
	var iPadding=objTable.cellPadding*2+parseInt(objTable.style.borderWidth);
/*
	for(var i=0;i<objTable.rows[0].cells.length;i++)
	{
		objTable.rows[0].cells[i].style.pixelWidth=objTable.rows[0].cells[i].offsetWidth-iPadding;
		objTable.rows[0].cells[i].style.pixelHeight=objTable.rows[0].cells[i].offsetHeight;
	}
*/
	var newDiv3=document.createElement("DIV");
	newDiv3.mergeAttributes(objDiv);
	newDiv3.style.overflow="hidden";
	newDiv3.style.zIndex=1000;

	objDiv.parentElement.insertBefore(newDiv3);
	var newDiv2=newDiv3.cloneNode(true);
	objDiv.parentElement.insertBefore(newDiv2);
	var newDiv1=newDiv3.cloneNode(true);
	objDiv.parentElement.insertBefore(newDiv1);
	var newTable1=document.createElement("TABLE");
	var newTBody=document.createElement("TBODY");
	newTable1.insertBefore(newTBody);
	var newTable2=newTable1.cloneNode(true);
	newTable1.mergeAttributes(objTable);
	newTable1.id="newTable1";
	newDiv1.insertBefore(newTable1);

	if(topRow){
		for(var i=0;i<topRow;i++)
		{
			var headRow=objTable.rows[i].cloneNode(true);
			newTable1.firstChild.insertBefore(headRow);
			for(var j=0;j<objTable.rows[i].cells.length;j++)
			{
				newTable1.rows[i].cells[j].style.pixelWidth=objTable.rows[i].cells[j].offsetWidth-iPadding;
//				objTable.rows[i].cells[j].style.pixelHeight=objTable.rows[i].cells[j].offsetHeight;
			}
			//objTable.rows[i].cells[0].style.pixelHeight=objTable.rows[i].cells[0].offsetHeight;
		}
		newTable1.style.pixelWidth=objTable.offsetWidth;
		newDiv1.style.pixelWidth=objDiv.clientWidth;
		newDiv1.style.pixelHeight=newTable1.offsetHeight+2;
		objDiv.attachEvent("onscroll",divScroll1);
	}
	if(leftRow){
		var newTable3=newTable2.cloneNode(true);
		newTable2.mergeAttributes(objTable);
		newTable3.mergeAttributes(objTable);
		newTable2.id="newTable2";
		newDiv2.insertBefore(newTable2);
		newDiv3.insertBefore(newTable3);
		var newTableWidth;
		var cols=new Array(objTable.rows.length);
		for(var i=0;i<objTable.rows.length;i++)
			cols[i]=leftRow;
		for(var i=0;i<objTable.rows.length;i++)
		{
			newRow=newTable2.insertRow();
			newRow.mergeAttributes(objTable.rows[i]);
			newTableWidth=0;
			for(var j=0;j<cols[i];j++)
			{
				var objCell=objTable.rows[i].cells[j].cloneNode(true);
				var rowSpan=objTable.rows[i].cells[j].rowSpan;
				if(rowSpan>1){
					for(var k=1;k<rowSpan;k++)
						cols[i+k]-=1;}
				newTableWidth+=objTable.rows[i].cells[j].offsetWidth;
				newRow.insertBefore(objCell);
				j+=objTable.rows[i].cells[j].colSpan-1;
			}
		}
		newTable2.style.pixelWidth=newTableWidth ;
		newDiv2.style.pixelHeight=objDiv.clientHeight;
		newDiv2.style.pixelWdith=newTable2.offsetWidth ;
		objDiv.attachEvent("onscroll",divScroll2);
	}
	if(topRow && leftRow){
		for(var i=0;i<topRow;i++){
			var headRow=newTable2.rows[i].cloneNode(true);
			newTable3.firstChild.insertBefore(headRow);
		}
		newTable3.style.pixelWidth=newTable2.offsetWidth;
		newDiv3.style.zIndex=1100;
		newDiv3.style.pixelWidth=newTable3.offsetWidth+2;
		newDiv3.style.pixelHeight=newTable3.offsetHeight+2;
		objDiv.attachEvent("onscroll",divScroll3);
	}
}

function divScroll1()
{
	newTable1.style.pixelLeft=-divMaster.scrollLeft;
}
function divScroll2()
{
	newTable2.style.pixelTop=-divMaster.scrollTop;
}
function divScroll3()
{
	newTable1.style.pixelLeft=-divMaster.scrollLeft;
	newTable2.style.pixelTop=-divMaster.scrollTop;
}


function CloneTableRow(objTable,iStartRow,iRows)
{	
	var objDiv=objTable.parentElement;
	var iPadding=objTable.cellPadding*2+parseInt(objTable.style.borderWidth);
	var newDiv1=document.createElement("DIV");
	newDiv1.mergeAttributes(objDiv);
	newDiv1.style.overflow="hidden";
	newDiv1.style.zIndex=1000;
	if (!objTable.assignedPixel)
	{
		
		for(var i=0;i<objTable.rows[0].cells.length;i++)
		{
			objTable.rows[0].cells[i].style.pixelWidth=objTable.rows[0].cells[i].offsetWidth-iPadding;
			objTable.rows[0].cells[i].style.pixelHeight=objTable.rows[0].cells[i].offsetHeight;
		}
		
		objTable.assignedPixel=true;
	}
	objDiv.parentElement.insertBefore(newDiv1);
	var newTable1=document.createElement("TABLE");
	var newTBody=document.createElement("TBODY");
	newTable1.insertBefore(newTBody);
	newTable1.mergeAttributes(objTable);
	newDiv1.insertBefore(newTable1);
	newDiv1.style.pixelWidth=objDiv.clientWidth;
	
	var iEndRow=iStartRow + iRows - 1;
	var sTemp="";
	for(var i=iStartRow;i<=iEndRow;i++)
		{
			var headRow=objTable.rows[i].cloneNode(true);
			newTable1.firstChild.insertBefore(headRow);
			
			for(var j=0;j<objTable.rows[i].cells.length;j++)
			{
				//newTable1.rows[i-iStartRow].cells[j].style.pixelWidth=objTable.rows[i].cells[j].offsetWidth - 2; // - iPadding
				newTable1.rows[i-iStartRow].cells[j].style.pixelWidth = parseInt(objTable.rows[i].cells[j].clientWidth) - parseInt(objTable.cellPadding)*2;
				//newTable1.rows[i-iStartRow].cells[j].mergeAttributes(objTable.rows[i].cells[j]);
				
//				if (newTable1.rows[i-iStartRow].cells[j].offsetWidth!=objTable.rows[i].cells[j].offsetWidth)
//				sTemp=sTemp+j+":("+newTable1.rows[i-iStartRow].cells[j].offsetWidth+"/"+
//					(objTable.rows[i].cells[j].offsetWidth) +")";
					
			}
			j=0;
			newTable1.rows[i-iStartRow].cells[j].style.pixelHeight=objTable.rows[i].cells[j].offsetHeight;
			
		}
	
	newTable1.style.pixelLeft=objTable.style.pixelLeft;
	newTable1.style.pixelTop=0;
	newTable1.style.position="absolute";
	newTable1.style.pixelWidth=objTable.offsetWidth;
	newDiv1.style.pixelHeight=newTable1.offsetHeight - 1;
	return newDiv1;

}


function fixedTableHead(objTable,headerNum,footerNum)
{
	var objDiv=objTable.parentElement;
	var iScroll=0;
	if(objDiv.scrollHeight<objDiv.offsetHeight)
		return false;
	if (headerNum)
	{
		var divHead=CloneTableRow(objTable,0,headerNum);
		divHead.style.pixelTop=objDiv.style.pixelTop;
		divHead.style.pixelLeft=objDiv.style.pixelLeft;
		divHead.children(0).id=objDiv.id+"_header";
	}
	if(footerNum){
		var divFooter=CloneTableRow(objTable,objTable.rows.length - footerNum,footerNum);
		divFooter.style.pixelTop=objDiv.offsetTop + objDiv.clientHeight + objDiv.clientTop - divFooter.offsetHeight + 1;
		divFooter.style.pixelLeft=objDiv.style.pixelLeft;
		var objFooter=divFooter.children(0);
		objFooter.id=objDiv.id+"_footer";
		objFooter.style.position="absolute";
		objFooter.style.pixelTop=-1;
	}
	objDiv.onscroll=divHScroll;
}

function fixedTableHead1(objTable,headerNum,footerNum)
{
	var objDiv=objTable.parentElement;
	var iScroll=0;
	var newDiv1=document.createElement("DIV");
	
	if(objDiv.scrollHeight<objDiv.offsetHeight)
		return false;
	if (headerNum)
	{
		var divHead=objTable.rows[1].cloneNode(true);
		divHead.style.pixelTop=objDiv.style.pixelTop;
		divHead.style.pixelLeft=objDiv.style.pixelLeft;
		divHead.children(0).id=objDiv.id+"_header";
	}
	if(footerNum){
		var divFooter= objTable.rows[objTable.rows.length - footerNum].cloneNode(true);
		divFooter.style.pixelTop=objDiv.offsetTop + objDiv.clientHeight + objDiv.clientTop - divFooter.offsetHeight + 1;
		divFooter.style.pixelLeft=objDiv.style.pixelLeft;
		var objFooter=divFooter.children(0);
		objFooter.id=objDiv.id+"_footer";
		objFooter.style.position="absolute";
		objFooter.style.pixelTop=-1;
	}
	objDiv.onscroll=divHScroll;
}

function divHScroll()
{
	var e=window.event.srcElement;

	if (document.all(e.id+"_header"))
	{
		document.all(e.id+"_header").style.pixelLeft=-e.scrollLeft;
	}
	if (document.all(e.id+"_footer"))
	{
		document.all(e.id+"_footer").style.pixelLeft=-e.scrollLeft;
	}
}

function findControlById(sID)
{
	var f=document.forms[0];
	for (var i=0;i<f.elements.length;i++)
	{
		if (f.elements[i].id && f.elements[i].id.indexOf(sID)>0) return f.elements[i];
	}
}

function getRequest(sKey)
{
	var iStart=0,iEnd=0;
	var s=""+window.location.search;
	iStart=s.indexOf(sKey+"=");
	if (iStart>0) 
	{
		iEnd=s.indexOf("&",iStart+1);
		if (iEnd<0) iEnd=s.length;
		return s.substring(iStart+sKey.length + 1,iEnd);
	}
	return null;
}

function getDate(oD,bFlag) {
	var st="";
	if (oD.type=="text")	st=oD.value;
	if (st.length==8&&!isNaN(st)) st=st.substr(0,4)+"/"+st.substr(4,2)+"/"+st.substr(6,2);
	var sDate = showModalDialog(LIBRARY_PATH+"fSysCalendar.htm",st,
		"dialogWidth=206px;dialogHeight=206px;center=yes;help=no;status=no");
	if (sDate&&sDate.length>0) 
	{
		if (oD.type=="text")	oD.value=sDate;
		return sDate;
	}
	else return null;
}
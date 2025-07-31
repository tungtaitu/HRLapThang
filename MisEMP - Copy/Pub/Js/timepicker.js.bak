//	Written by Tan Ling wee
//	on 19 June 2005
//	email :	info@sparrowscripts.com
//      url : www.sparrowscripts.com

/////////////////////////////////Here is the fix from Sihui Wu 
// More Fixes by Mahesh, 5/2007
// - Handles a wider variety of short time formats
// - When multiple time elements are used, ensures that if the picker button 
//   is pressed, the previously pressed button is unpressed.
// - Widget shows time closest to time in textbox element
// - Widget is modal and throws a curtain over other screen elements
// - If am/pm is not specified, all hours within 9-11 inclusive are considered to be AM, and the rest PM.
//   This is consistent with usage of time in speech, for business hours.
// - Fixed so that pressing ESC closes the timepicker
// - Fixed script so that it works with Yahoo! maps AJAX API v.3.4. Variable dom inteferes with the Yahoo! script, so
// - has been renamed to docGeid.

//By Jeffrey 2007-08-30 http://www.darkthread.net
//Adjust the script to be a "textbox extender" and using 24H style
// - Add AFA_TIMEPICKER_IMGPATH to specify cust images path
// - Add to24H function to convert am/pm to 24 hours format
// - setTimePicker(), validateDatePicker2() add am/pm to 24H conversion
// - Add afa_TimePickerOnBlur, afa_TimePickerOnClick, afa_ExtendTimePicker

  // modify image path to suit your application.
  // var imagePath='widgets/timePicker/images/';
  var imagePath="../../App_Themes/pwfBody/images/datetimeImg/";
  //2007-08-30 by Jeffrey, accept custom image path
  if (AFA_TIMEPICKER_IMGPATH) {
  	imagePath = AFA_TIMEPICKER_IMGPATH;
  }

  var ie=document.all;
  var docGeid=document.getElementById;
  var ns4=document.layers;
  var bShow=false;
  var textCtl; // "bogus" form field?? what is this for???
  
  var debugging=false; // for use in logging.

  function to24H(t) {
  	//by Jeffrey 2007-08-30
  	//Convert to 24H
  	if (t.indexOf("am")==-1 && t.indexOf("pm")==-1) return t;
  	var p = t.substring(0, t.length-3).split(':');
  	if (p[0] == 12) p[0] = 0;
  	//if (t.indexOf("pm")>0) p[0]=p[0]*1 + 12; //to 24H
	if (t.indexOf("pm")>0) p[0]=p[0]*1 + 0; //to 24H
  	if (p[0]<10) p[0]='0'+p[0];
  	return p[0] + ':' + p[1];
  }


  //function tpDId(eleId) { return document.getElementById(eleId); }

  // t is an element.
  function setTimePicker(t) {
    textCtl.value=to24H(t);
	//textCtl.value = t;
    // if (typeof(validateTime)!="undefined")
    //  validateTime(textCtl); // try to get rid of this... this is reference OUTSIDE of the timepicker!
    //validateDatePicker(textCtl);
    closeTimePicker();
  }

  /* nearestTime: find the nearest time within 15 mins.
     assumption: time is formatted as xx:xx am/pm or xx:xx:xx am/pm, and is a valid time.
  */
  function nearestTime(n) {
    if (debugging) logger("nearestTime");
    var t = new Object();
    t.value = n;
    // t is not a screen object, so call vDP2. vDP itself expects to be given a reference to a textbox.
    validateDatePicker2(t);
    
    // time is now validated; should be in xx:xx am/pm format.
    var arr = t.value.split(":");
    var a2=arr[1].split(" ");
    var mn=parseInt(a2[0],10);
    var ampm=a2[1];
    
    // get nearest minute boundary, within 15 mins.
    var nMin= parseInt((mn+7)/15, 10)*15;
    
    return arr[0]+":"+padZero(nMin)+" "+ampm;
    
    
  }


  /*
    mode: am or pm
    tm: time selected, must be properly formatted: xx:xx am/pm, e.g. 12:30 pm
    
    if tm is provided, then mode is ignored.
  */
  function refreshTimePicker(mode, tm) {
    // was a selected time provided?
    if (tm===undefined) { // is ===undefined an error, or is this correct javascript syntax?
        if (mode==0)
          { 
            suffix="am"; 
          }
        else
          { 
            suffix="pm"; 
          }
    } else {
      tm = nearestTime(tm); // get time to nearest 15 min. interval.
      suffix=tm.split(" ")[1];
      if (suffix=="am") 
        mode = 0; 
      else 
        mode = 1;
    }

    if (mode==0) {
        document.getElementById("iconAM").src= imagePath + "am1.gif";
        document.getElementById("iconPM").src= imagePath + "pm2.gif";
    } else {
        document.getElementById("iconAM").src= imagePath + "am2.gif";
        document.getElementById("iconPM").src= imagePath + "pm1.gif";
    }

    // sHTML = "<table><tr><td><table cellpadding=3 cellspacing=0 bgcolor='#f0f0f0'>";
    sHTML = "<table><tr><td><table cellpadding=3 cellspacing=0 bgcolor='#FFFFFF'>";
    for (i=0;i<12;i++) {

      sHTML+="<tr align=right style='font-family:verdana;font-size:11px;color:#000000;'>";

      if (i==0) {
		if(suffix=="pm"){
			hr = 12;
		}else{
			hr = 0;
		}
		//hr = 0;
      }
      else {
        hr=i;
		if(suffix=="pm"){
			hr+=12;
		}
      }  

	  //suffix="";
      for (j=0;j<4;j++) {
        var thisTime=hr+":"+padZero(j*15)+" " + suffix;
        var bgcolor = "";
        if (thisTime==tm) {bgcolor="bgcolor='#F7C8E3'";}
        //sHTML+="<td " + bgcolor + "width=57 style='cursor:hand;font-family:verdana;font-size:11px;' onmouseover='this.style.backgroundColor=\"silver\"' onmouseout='this.style.backgroundColor=\"\"' onclick='setTimePicker(\""+ hr + ":" + padZero(j*15) + "&nbsp;" + suffix 
        //+ "\")'><a style='text-decoration:none;color:#000000' href='javascript:setTimePicker(\""+ hr + ":" + padZero(j*15) + "&nbsp;" + suffix + "\")'>" + hr + ":"+padZero(j*15) +"&nbsp;"+ "<font color=\"#808080\">" + suffix + "&nbsp;</font></a></td>";
		
        sHTML+="<td " + bgcolor + "width=57 style='cursor:hand;font-family:verdana;font-size:11px;' onmouseover='this.style.backgroundColor=\"silver\"' onmouseout='this.style.backgroundColor=\"\"' onclick='setTimePicker(\""+ hr + ":" + padZero(j*15) + "&nbsp;" + suffix 
        + "\")'><a style='text-decoration:none;color:#000000' href='javascript:setTimePicker(\""+ hr + ":" + padZero(j*15) + "&nbsp;" + suffix + "\")'>" + hr + ":"+padZero(j*15) +"&nbsp;"+ "<font color=\"#808080\">" + "" + "&nbsp;</font></a></td>";		

      }

      sHTML+="</tr>";
    }
    sHTML += "</table></td></tr></table>";
    document.getElementById("timePickerContent").innerHTML = sHTML;
	
  }

  if (docGeid){
    document.write ("<div id='timepicker' style='z-index:9;position:absolute;visibility:hidden;'><table style='border-width:1px;border-style:solid;border-color:gray;' bgcolor='#ffffff' cellpadding=0><tr bgcolor='gray'  ><td><table cellpadding=0 cellspacing=0 width='100%' ><tr valign=bottom height=21><td style='font-family:verdana;font-size:11px;color:#ffffff;padding:3px' valign=center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td><img id='iconAM' src='" + imagePath + "am1.gif' onclick='document.getElementById(\"iconAM\").src=\"" + imagePath + "am1.gif\";document.getElementById(\"iconPM\").src=\"" + imagePath + "pm2.gif\";refreshTimePicker(0)' style='cursor:hand'></td><td><img id='iconPM' src='" + imagePath + "pm2.gif' onclick='document.getElementById(\"iconAM\").src=\"" + imagePath + "am2.gif\";document.getElementById(\"iconPM\").src=\"" + imagePath + "pm1.gif\";refreshTimePicker(1)' style='cursor:hand'></td><td align=right valign=center style='padding-top:7px'>&nbsp;<img onclick='closeTimePicker()' src='" + imagePath + "close.gif'  STYLE='cursor:hand'>&nbsp;</td></tr></table></td></tr><tr><td colspan=2><span id='timePickerContent'></span></td></tr></table></div>");
    refreshTimePicker(0);
  }

  var crossobj=(docGeid)?document.getElementById("timepicker").style : ie? document.all.timepicker : document.timepicker;
  var currentCtl;


/*
  // capture window resize event
  function register(e) {
    curtain.setScreenerSize();
	  return true;
  }
*/

  // get absolute position of a control. use with overlays, dropdowns, etc.
  function getAbsPos(ctl) {
    var leftpos=0
    var toppos=0
    var aTag = ctl
    do {
      aTag = aTag.offsetParent;
      leftpos  += aTag.offsetLeft;
      toppos += aTag.offsetTop;
    } while(aTag.tagName!="BODY");
    
    var o= new Object();
    o.left=leftpos
    o.top = toppos;
    return o;
  }    

  
  // show time picker. ctl is the timepicker button (image) pressed. ctl2 is the text element to be populated with the time.
  function selectTime(ctl,ctl2) {
    if (debugging) logger("selectTime:"+ctl2.id);

    textCtl=ctl2;
    
    /* Modification below by MV, 5/2007. If you have multiple time pickers on a page, 
    and click one picker, and then click another one without closing the first, this modification
    ensures that the Button for the first picker reverts to unpressed. In the original script,
    if timePicker buttons were pressed one after the other, they would all show up as pressed. */
    

    
    if ((currentCtl != ctl) && (currentCtl != null)) {// not the same
        currentCtl.src=imagePath + "timepicker.gif"; // prev button in released state
    }
    currentCtl = ctl;
    currentCtl.src=imagePath + "timepicker2.gif"; // curr button pressed state
	
    // let the timepicker show a time as close to the current choice as possible, if the time
    // in the textbox is valid.
    if (ctl2.value!="") {
    		//by Jeffre, 2007-08-30 convert the 24H format to am/pm
    		var t = to24H(ctl2.value);
    		if (t.match(/^[0-2][0-9]:[0-5][0-9]$/)) {
    			var p = t.split(':');
    			if (p[0]>=0 && p[0]<=11)
    				refreshTimePicker(0);
    			else
    				refreshTimePicker(1);
    		} else {
	        var res=validateDatePicker2(ctl2);
	        if (res)
	            refreshTimePicker(0, ctl2.value);
	        else
	            refreshTimePicker(0);
        }
    }
    
    aPos = getAbsPos(ctl);
    
    crossobj.left =  ctl.offsetLeft  + aPos.left;
    crossobj.top = ctl.offsetTop +  aPos.top + ctl.offsetHeight +  2 
    crossobj.visibility=(docGeid||ie)? "visible" : "show"
    hideElement( 'SELECT', document.getElementById("timepicker") );
    hideElement( 'APPLET', document.getElementById("timepicker") );
    

    
    // make the time picker modal.
    curtain.show();
    
    bShow = true;
  }

  // hides <select> and <applet> objects (for IE only)
  function hideElement( elmID, overDiv ){
    if( ie ){
      for( i = 0; i < document.all.tags( elmID ).length; i++ ){
        obj = document.all.tags( elmID )[i];
        if( !obj || !obj.offsetParent ){
            continue;
        }
          // Find the element's offsetTop and offsetLeft relative to the BODY tag.
          objLeft   = obj.offsetLeft;
          objTop    = obj.offsetTop;
          objParent = obj.offsetParent;
          while( objParent.tagName.toUpperCase() != "BODY" )
          {
          objLeft  += objParent.offsetLeft;
          objTop   += objParent.offsetTop;
          objParent = objParent.offsetParent;
          }
          objHeight = obj.offsetHeight;
          objWidth = obj.offsetWidth;
          if(( overDiv.offsetLeft + overDiv.offsetWidth ) <= objLeft );
          else if(( overDiv.offsetTop + overDiv.offsetHeight ) <= objTop );
          else if( overDiv.offsetTop >= ( objTop + objHeight + obj.height ));
          else if( overDiv.offsetLeft >= ( objLeft + objWidth ));
          else
          {
          obj.style.visibility = "hidden";
          }
      }
    }
  }
     
  //unhides <select> and <applet> objects (for IE only)
  function showElement( elmID ){
    if( ie ){
      for( i = 0; i < document.all.tags( elmID ).length; i++ ){
        obj = document.all.tags( elmID )[i];
        if( !obj || !obj.offsetParent ){
            continue;
        }
        obj.style.visibility = "";
      }
    }
  }

  function closeTimePicker() {
	//alert('close');
    bShow=false;
    crossobj.visibility="hidden"
    showElement( 'SELECT' );
    showElement( 'APPLET' );
    currentCtl.src=imagePath + "timepicker.gif"
    
    curtain.hide();
  }

/*
  document.onkeypress = function hideTimePicker1 (event) { 
    if (event.keyCode==27){
      //if (!bShow){
      if (bShow){
        closeTimePicker();
      }
    }
  }
  
  */
  document.onkeypress = function(e) {
    var keynum;
    if (window.event) // IE
      keynum=window.event.keyCode;
    else // Netscape/Firefox/Opera
      keynum = e.keyCode; 
      
    if (keynum == 27) 
      if (bShow)
        closeTimePicker();
  }
  //The onblur event for hooked textbox
  function afa_TimePickerOnBlur(evt) {
  	 var txt;
  	 if (window.event) 
  	 	txt = window.event.srcElement;
  	 else
  	 	txt = evt.target;
  	 validateDatePicker(txt);
  }
  //The onclick event form timepicker icon
  function afa_TimePickerOnClick(evt) {
  	 var img;
  	 if (window.event) 
  	 	img = window.event.srcElement;
  	 else
  	 	img = evt.target;
  	 selectTime(img, document.getElementById(img.targetId));
	 var intObjHeight = document.getElementById("timepicker").offsetHeight;
	 var intObjWidth  = document.getElementById("timepicker").offsetWidth;
	 var intTop		  = document.getElementById("timepicker").offsetTop;
	 var intLeft	  = document.getElementById("timepicker").offsetLeft;
	 var intWinHeight = document.body.clientHeight;
	 var intWinWidth  = document.body.clientWidth;
	 
	 //alert(intLeft);
	 //alert(img.offsetHeight);
	 if((intObjHeight+intTop)>(intWinHeight)){
		document.getElementById("timepicker").style.top = document.getElementById("timepicker").offsetTop - (document.getElementById("timepicker").offsetHeight+img.offsetHeight);
	 }
	 
	 //alert(intLeft);
	 if((intObjWidth+intLeft)>(intWinWidth)){
		document.getElementById("timepicker").style.left = ((intLeft - intObjWidth)-15);
	 }
  }  
  //Hooking the textbox to extend the timepicker function
 	function afa_ExtendTimePicker(txtId) {
		var txt = document.getElementById(txtId);
		if (!txt) {
			alert("Textbox["+txtId+"] not found!");
			return;
		}
		//Hook onblur
		txt.onblur = afa_TimePickerOnBlur;
		//Add picker icon
		var img = document.createElement("IMG");
		img.src = imagePath + "timepicker.gif";		
		img.style.cursor = "hand";
		img.width = 30;
		img.height = 20;
		img.setAttribute("ALT", "Pick a Time!");
		img.targetId = txt.id;
		img.onclick = afa_TimePickerOnClick;
		txt.parentNode.insertBefore(img, txt.nextSibling); 		
 	}

  function isDigit(c) {
    
    return ((c=='0')||(c=='1')||(c=='2')||(c=='3')||(c=='4')||(c=='5')||(c=='6')||(c=='7')||(c=='8')||(c=='9'))
  }

  function isNumeric(n) {
    
    num = parseInt(n,10);

    return !isNaN(num);
  }

  function padZero(n) {
    v="";
    if (n<10){ 
      return ('0'+n);
    }
    else
    {
      return n;
    }
  }

   // if the hour is between 9 to 11, assume it is AM. if it is between 12-8, assume it is PM
  function amOrPm(hr) {
    if ((parseInt(hr,10)>=9) && (parseInt(hr)<=11))
         return "am"
    else
        return "pm"
  }

    // validate whether the contents of a textbox represent a valid time.
   function validateDatePicker(ctl) {
        if (debugging) logger("validateDatePicker");
        var res=validateDatePicker2(ctl);
        if (res!=true)
            ctl.style.color="#FF0000";
        else
            ctl.style.color="#000000";
        return res;
   }

  // validate the time
  function validateDatePicker2(ctl) {
    if (debugging) logger("validateDatePicker2");
    t=ctl.value.toLowerCase();
    t=t.replace(" ","");
    t=t.replace(".",":");
    t=t.replace("-","");

    if ((isNumeric(t))&&(t.length==4))
    {
      t=t.charAt(0)+t.charAt(1)+":"+t.charAt(2)+t.charAt(3);
    }

    var t=new String(t);
    tl=t.length;

    if (tl==1 ) {
      if (isDigit(t)) {
        if (t=="0") 
            ctl.value="12:00 am";
        else
            ctl.value=t+":00" +amOrPm(t);
      }
      else {
        return false;
      }
    }
    else if (tl==2) {
      if (isNumeric(t)) {
        if (parseInt(t,10)<13){
          if (t.charAt(1)!=":") {
            ctl.value= t + ':00' + amOrPm(t);
          }
          else {
            if (t.charAt(0)=="0")
                ctl.value="12:00 am";
            else
                ctl.value= t + '00' + amOrPm(t);
          }
        }
        else if (parseInt(t,10)==24) {
          ctl.value= "12:00 am";
        }
        else if (parseInt(t,10)<24) {
          if (t.charAt(1)!=":") {
            ctl.value= (t-12) + ':00 pm';
          } 
          else {
            ctl.value= (t-12) + '00 pm';
          }
        }
        else if (parseInt(t,10)<=60) {
          ctl.value= '0:'+padZero(t)+' am';
        }
        else {
          ctl.value= '1:'+padZero(t%60)+' am';
        }
      }
      else
           {
        if ((t.charAt(0)==":")&&(isDigit(t.charAt(1)))) {
          ctl.value = "0:" + padZero(parseInt(t.charAt(1),10)) + " am";
        }
        else {
          return false;
        }
      }
    }
    else if (tl>=3) {

        // 3-digit time modification by MV, 5/2007
        if ((tl==3) && (!isNumeric(t))) return false;
        if ((tl==3) && (isNumeric(t))) {
            // time is in format, say 330, for 330 am or pm
            var tHour=t.charAt(0);
            var tMin=t.charAt(1)+t.charAt(2);
            hr=parseInt(tHour,10);
            mn=parseInt(tMin,10);
            if (isNaN(mn)) mn=0; // e.g. if "7qq" is entered, this becomes 7:00pm
            if ((mn<0) || (mn>59))
                return false;
            if (hr==0) {
                hr=12;
                mode="am";
            } else
                mode=amOrPm(tHour);            
            
            ctl.value=hr+":"+padZero(mn)+" "+mode;
            return true;
        }

      // now tl>3
      var arr = t.split(":");
      if (t.indexOf(":") > 0)
      {
        hr=parseInt(arr[0],10);
        mn=parseInt(arr[1],10);

        if (t.indexOf("pm")>0)
          mode="pm";
        else if (t.indexOf("am")>0)
          mode="am";
        else
          mode=amOrPm(hr);

        if (isNaN(hr)) {
          return false;
          hr=0;
        } else {
          if (hr>24) {
            return false;
          }
          else if (hr==24) {
            mode="am";
            hr=0;
          }
          else if (hr>12) {
            mode="pm";
            hr-=12;
          }
        }
      
        if (isNaN(mn)) {
          mn=0;
        }
        else {
          if (mn>60) {
            mn=mn%60;
            hr+=1;
          }
        }
      } else {

        hr=parseInt(arr[0],10);

        if (isNaN(hr)) {
          return false;
          // hr=0;
        } else {
          if (hr>24) {
            return false;
          }
          else if (hr==24) {
            mode="am";
            hr=0;
          }
          else if (hr>12) {
            mode="pm";
            hr-=12;
          }
        }

        mn = 0;
      }
      
      if (hr==24) {
        hr=0;
        mode="am";
      }
      ctl.value=hr+":"+padZero(mn)+" "+mode;
    }
    //by Jeffrey @ 2007-08-30 Convert it to 24H
    ctl.value = to24H(ctl.value);
    return true;
  }

/**********************************************************************************************
  Curtain object & methods. Throw a translucent curtain on screen objects.
 **********************************************************************************************/

  var Curtain = function(id){
    //document.write("<div id='"+id+"' style='z-index:8;visibility:hidden;background: #333; opacity: 0.10; filter: alpha(opacity=10); position: absolute; top:0; left:0; width: 100%; height: 100%;'>&nbsp;</div>");
	document.write("<div id='"+id+"' style='z-index:8;visibility:hidden;background: #333; opacity: 0; filter: alpha(opacity=0); position: absolute; top:0; left:0; width: 100%; height: 100%;'>&nbsp;</div>");
    this.id=id;
  };
  
  Curtain.prototype.show = function () {
    this._scrn().style.visibility="visible";
    this._autoResize();
  }

  Curtain.prototype.hide = function() {
    this._scrn().style.visibility="hidden";
  }
  
  // Modal overlay screen element
  Curtain.prototype._scrn = function() {
      return document.getElementById(this.id);
/*      var ie=document.all;
      return (document.getElementById)?
        document.getElementById(this.id) : 
        ie? document.all.screener : document.screener;
*/        
  }
 

  // set gray overlay screener size to document size, including scrollbars
  Curtain.prototype._autoResize = function() {
    // try to set to 0, so that the overlay does not influence the document size. this may or may not work...
    // seems to work in firefox, but the height does not readjust in IE6.
    this._scrn().style.width=0;
    this._scrn().style.height=0;
    if (ie) {
        this._scrn().style.width=document.body.scrollWidth;
        this._scrn().style.height=document.body.scrollHeight;
    } else {
        this._scrn().style.width=document.body.scrollWidth;
        this._scrn().style.height=document.documentElement.scrollHeight;
    }
  }

  // capture window onresize, so that curtain can be resized. 
  window.onresize = function() { curtain._autoResize(); }
  
  curtain = new Curtain();

 
/**********************************************************************************************/

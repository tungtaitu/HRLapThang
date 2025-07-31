function fn_openWin(s_url,n_width,n_height,s_winName,b_scroll){
	if(!s_winName){
		s_winName = "";
	}
	if(!b_scroll){
		b_scroll=0;
	}else{
		b_scroll=1;
	}
	var s_winFeatures="top="+((screen.height-n_height)/2-30)+",left="+((screen.width-n_width)/2-5)+",height="+n_height+",width="+n_width+",scrollbars="+b_scroll+",status=no,toolbar=no,menubar=no,location=no,resizable=0";
	return(window.open(s_url,s_winName,s_winFeatures));
}

function fn_openWinFull(s_url,s_winName,b_scroll){
	if(!s_winName){
		s_winName="";
	}
	if(!b_scroll){
		b_scroll=0;
	}else{
		b_scroll=1;
	}
	var s_winFeatures="top=0,left=0,height="+(screen.availHeight-30)+",width="+(screen.availWidth-10)+",scrollbars="+b_scroll+",status=0,toolbar=no,menubar=0,location=no,resizable=1";
	return(window.open(s_url,s_winName,s_winFeatures));
}

function fn_openDialog(s_url,n_width,n_height,b_scroll,o_arg){
	if(!o_arg){
		o_arg="";
	}
	if(!b_scroll){
		b_scroll=0;
	}
	var winFeatures="center=1;dialogHeight="+(parseInt(n_height)+29)+"px;dialogWidth="+(parseInt(n_width)+10)+"px;status=no;help=no;scroll="+b_scroll;
	return(window.showModalDialog(s_url,o_arg,winFeatures));
}

function fn_openDialogFull(s_url,b_scroll,o_arg){
	if(!o_arg){
		o_arg="";
	}
	if(!b_scroll){
		b_scroll=0;
	}
	var winFeatures="center=1;dialogHeight="+(screen.availHeight-0)+"px;dialogWidth="+(screen.availWidth-0)+"px;status=no;help=no;scroll="+b_scroll;
	return(window.showModalDialog(s_url,o_arg,winFeatures));
}

/* show groupuser list detail */
function show_groupuser(objSender){
    var trDetail = $(objSender).parent().parent().next();
    var tdDetail = $(objSender).parent().parent().next().find("td:eq(0)");
    //loading data

    if(jQuery.trim($(tdDetail).html())==""){
        var GroupID = $(objSender).parent().find(":hidden:eq(0)").val();             
        getDetail(tdDetail,GroupID);
    }
    //toggle power list detail show/hide
    if(jQuery.trim($(trDetail).css("display"))=="none"){
        $(trDetail).show();
    }else{
        $(trDetail).hide();
    }
    //release ram        
    trDetail = null;
    tdDetail = null;
}

function getDetail(objSender,GroupID){
    var objServerVar;
    $(objSender).html("<div style='color:gray;text-align:center'>Loading data, please wait ....<br><img src='../../App_Themes/pwfBody/images/AjaxLoadingBar1.gif'></div>");
	$.ajax(
		{
			url:"../../Pub/RemoteMethods/GroupUser.aspx",
			type:"post",
			data:"EventName=getDetail&GroupID="+GroupID,
			dataType:"text",
			timeout:100000,
			error:function(err){
				objServerVar = {
					IsOK:false,
					ServerMsg : "遠端連線發生錯誤或連線逾時，請再試一次，如果持續發生請連絡相關人員！！"
				}
				$(objSender).html("");
				$(objSender).parent().hide();
			},
			success:function(strServerVar){
				eval("objServerVar ="+strServerVar);
			},
			complete:function(){
				if(objServerVar.IsOK){
				    //alert(objServerVar.Result);
					$(objSender).html( objServerVar.Result );
				}else{
				    alert(objServerVar.ServerMsg);
				}
			}
			
		}
	);        
}
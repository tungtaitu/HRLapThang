<%@LANGUAGE=VBSCRIPT CODEPAGE=65001%> 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1">   
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

    <link href="Pub/Tab/Css/IndexStyle.css" rel="stylesheet" type="text/css" />
    <script src="Pub/Tab/Js/jquery.js" type="text/javascript"></script>
    <script src="Pub/Tab/Js/jquery-ui/js/jquery-ui.js" type="text/javascript"></script>
    <link href="Pub/Tab/Js/jquery-ui/css/redmond/jquery-ui.custom.css" rel="stylesheet"  type="text/css" />
    <link href="Pub/Css/Default.css" rel="stylesheet" type="text/css" />
    <link href="Pub/Css/PccStyles.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="Pub/JS/fn_btnOver_v02.js"></script>
    <script type="text/javascript" src="Pub/JS/fn_imgOver.js"></script>
    <script type="text/javascript" src="Pub/JS/fn_open.js"></script>

    <script type="text/javascript" src="Pub/JS/Common.js"></script>
	<style type="text/css">
	
	    /*#tabs { margin-top: 0px; background-color:#2e5cb7 }*/
	    #tabs { margin-top: 0px; background: url(Images/TabImg/header_frame.jpg) repeat ;  border:0px; }
	    #tabs ul {  height:24px; }		
	    #tabs li .ui-icon-close { float: left; margin:0; cursor: pointer; }
	    #add_tab { cursor: pointer; }
	    #tabs-1_title{ width:auto; margin-left:4px;}
	    #btn_MenuHide{ width:auto; margin-left:2px;}
	    #btn_MenuShow{ width:auto; margin-left:2px;}
	   
	    .ui-tabs .ui-tabs-panel { display: block; border-width: 0; padding: 0; background: none; }
	    
	    .ui-tabs { position: relative; padding: 0;  }
        .ui-tabs .ui-tabs-nav { margin: 0; padding: 0; }
        .ui-tabs .ui-tabs-nav li { list-style: none; float: left; position: relative; top: 0px; margin: 0 .2em 1px 0; border-bottom: 0 !important; padding: 0; white-space: nowrap; }
        .ui-tabs .ui-tabs-nav li a { float: left; padding: 6px 1px 0px 6px; text-decoration: none; }
        .ui-tabs .ui-tabs-nav li.ui-tabs-selected { margin-bottom: 0; padding-bottom: 1px; }
        .ui-tabs .ui-tabs-nav li.ui-tabs-selected a, .ui-tabs .ui-tabs-nav li.ui-state-disabled a, .ui-tabs .ui-tabs-nav li.ui-state-processing a { cursor: text; }
        .ui-tabs .ui-tabs-nav li a, .ui-tabs.ui-tabs-collapsible .ui-tabs-nav li.ui-tabs-selected a { cursor: pointer; } 
        .ui-tabs .ui-tabs-hide { display: none !important; }
       
        .ui-widget-content { border: 0px;  }
        .ui-widget-header { height:24px; background: url(Images/TabImg/header_frame.jpg) repeat ;  border:0px; }      

	</style>
	
	<script type="text/javascript">
        $(function () {
            var tab_counter = 2;
            var $tabs = $("#tabs").tabs({
                tabTemplate: "<li><a href='#{href}' class='tabTitle'>#{label}</a> <span class='ui-icon ui-icon-close'>Remove Tab</span></li>",
                add: function (event, ui) {
                    var tab_content = $("#hidCurrentLink").val();
                    var tab_id = "tab_counter_" + tab_counter;
                    $(ui.panel).append('<iframe id="' + tab_id + '" style="width:100%;height:100%;"  src="' + tab_content + '" frameborder="0"></iframe>');
                }
            });

            function addTab(mnuName) {
                $tabs.tabs("add", "#tabs-" + tab_counter, mnuName);
            }

            $("#tabs span.ui-icon-close").live("click", function () {
                //-2 vi co them 2 nut dong va mo menu(btn_MenuHide va btn_MenuShow)
                var index = $("li", $tabs).index($(this).parent())-2;
                //alert(index);
                $tabs.tabs("remove", index);
            });

            $(".menu_Link").click(function () {
                var mnuName = $(this).attr("mnuName");
                var mnuLink = $(this).attr("mnuLink");

                $("#hidCurrentLink").val(mnuLink);

                var i = FindTabsIndex(mnuName);
                if (i == -1) {
                    addTab(mnuName);
                    $("#tabs").tabs("select", $(".tabTitle").length);
                }
                else {
                    $("#tabs").tabs("select", i + 1);
                }

                $("#hidCurrentTab").val("tab_counter_" + tab_counter);

                menuSlide();
                tab_counter++; // Tang id Tab lên 1 đơn vị

            });

            function FindTabsIndex(tabTitle) {
                var i = -1;
                $(".tabTitle").each(function (index) {
                    if ($(this).html() == tabTitle) {
                        i = index;
                    }
                });
                return i;
            }
        });
	</script>
	
	<script type="text/javascript">
        $(document).ready(function () {
            $("#hidCurrentTab").val("iframeContent");
            LoadDivMain();            
        });

        $(window).resize(function () {           
            var mainContent_height = $(window).height() - $("#pnTopHeader").height() - $("#pnToolBar").height() - $("#pnFooter").height() - 12;
            $("#tdFrame").height(mainContent_height);
        });

        function LoadDivMain() {
            var mainContent_height = $(window).height() - $("#pnTopHeader").height() - $("#pnToolBar").height() - $("#pnFooter").height() - 12;
            
            $("#tdFrame").height(mainContent_height);
            $("#iframeContent").height($("#tdFrame").height() - 26);
			$("#tableMain").height(mainContent_height);
        }



        function menuSlide() {
            var mainContent_height = $(window).height() - $("#pnTopHeader").height() - $("#pnToolBar").height() - $("#pnFooter").height() - 12;
            var firstpane_height = $("#firstpane").height();

            var currentTab = $("#hidCurrentTab").val();

            if (firstpane_height < mainContent_height) {
                $("#" + currentTab).height(mainContent_height - 26);
                $("#tdFrame").height(mainContent_height);
                $("#tableMain").height(mainContent_height);
            }
            else {
                $("#" + currentTab).height(firstpane_height - 26);
                $("#tdFrame").height(firstpane_height);
                $("#tableMain").height(firstpane_height);
            }
        }

        function changeFrame(link) {
            addTab();
        }
    </script>
	
	<script type="text/javascript" >
        
        var o_displayed = null;
        function fnHead(n_index) 
        {
            var mainContent_height = "";
            if (n_index == -1) {
                mainContent_height = $(window).height() - $("#pnToolBar").height() - $("#pnFooter").height()-10;
				
                pnTopHeader.style.display = "none";
                pnShowHeader.style.display = "";
                pnHideHeader.style.display = "none";                
				tableMain.style.height =mainContent_height;	
				
				//alert($("#tabs").height());
				//alert($("#tabs_1").height());
				//alert($("#tabs_1").height()+2);
				//$("#tabs_1").height()=mainContent_height-30;
				//tabs_1.style.height=615;
				//alert($("#tabs_1").height()+2);
				//iframeContent.style.height=$("#tabs").height()-$("#tabs_1").height();
				
				
            } else {
                mainContent_height = $(window).height() - $("#pnTopHeader").height() - $("#pnToolBar").height() - $("#pnFooter").height()-12;
                pnTopHeader.style.display = "";
                pnShowHeader.style.display = "none";
                pnHideHeader.style.display = "";
                tableMain.style.height = mainContent_height;
				
				//alert(tdFrame.style.height);
				//alert(tabs.style.height);
            }
        }

        function fn_switchVisible(obj) {
            obj.style.display = (obj.style.display == "none" ? "" : "none");
        }
        function _onPrevious() {
            history.back();
        }
        function _onNext() {
            history.forward();
        }
	</script>
    
</head>
<body onload="fn_MenuHide()">
    <form id="System" method="post">
        <input id="hidCurrentLink" type="hidden" />
        <input id="hidCurrentTab" type="hidden" />   
        <div id="pnTopHeader">
            <table cellpadding="0" cellspacing="0" style="width: 100%;height:70px">
                <tr>
                    <td style="background: url(App_Themes/banner.png) 0 100%;" align="left"></td>                    
                </tr>
            </table>
        </div>  
		
        <div id="pnToolBar">               
            <div id="pnHideHeader" onclick="fnHead(-1)" style="width:30px; cursor: hand;display: none;">
                <img src="Images/TabImg/arrow_up.gif" alt="" /> 
            </div>
            <div id="pnShowHeader" onclick="fnHead(0)" style="width:30px; cursor: hand; display: none;">
                <img src="Images/TabImg/arrow_down.gif" alt="" />
            </div>            
            <div id="pnHome" style="width:10%">
                <a href="main.asp" target="main">
                    <img src="Images/TabImg/home.gif"  style="height:18px;" alt="Home" />&nbsp;訊息頁 
                </a>
            </div>
			<div id="pnLogOut"  style="width:20%">
				<a href="default.asp?logout=Y" target="_top">
					<img style=" text-align: right; vertical-align: middle;height:18px;" src="Images/TabImg/LogOut.gif" />&nbsp;登出系統
				</a>
			</div>
			<div id="pnRefresh"  style="width:20%">
                <a id="A1" href="javascript:void(0)" onclick="_onPrevious()" style="color:Gray;">
					<img src="Images/TabImg/previousover.gif"/>上一頁
                </a>
				&nbsp;&nbsp;
                <a id="A4" href="javascript:void(0)" onclick="_onNext()" style="color:Gray;">
					<img src="Images/TabImg/NextOver.gif" />下一頁
                </a>  
				&nbsp;&nbsp;
                <a id="ClickRefresh" href="javascript:void(0)"  style="color:Gray;" >
					<img src="Images/TabImg/refreshover.gif" />重新整理
				</a>
            </div> 
			<div id="pnUser"  style="width:20%">
				<span>登入者 : <%=Session("netuser")%>(<%=REMOTE_IP%>)</span>
			</div>
            
        </div>
            
        <table id="tableMain" cellpadding="0" cellspacing="2"  style="width:100%">				
			<tr>
				<td id="tdLeftMenu"  valign="top" align="left" class="main_td"> 
					<div id="firstpane" class="menu_list" style="width:100%;height:100%;">
                        <iframe name="contents" id="textareaCode" target="main" style="width:100%;height:100%;" src="" frameborder="0"></iframe>    
                    </div>                        
				</td>				
				<td id="tdFrame"  valign="top" align="left" style="height:100%">
                    <div id="tabs">
                        <ul>
                            <li id="btn_MenuHide"><img src="App_Themes/pwfBody/images/menuHide_0.gif" alt="" /></li>
                            <li id="btn_MenuShow" style="display:none"><img ID="imgshow" src="App_Themes/pwfBody/images/menuShow_0.gif" alt="" /></li> 
		                </ul>                        
		                <div id="tabs-1">
                            <iframe name="main" id="iframeContent" style="width:100%;height:100%;" src="main.asp" frameborder="0"></iframe>	                    
                        </div>
	                </div>
                </td>				
			</tr>
		</table>
        
       
        <div id="pnFooter">
            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height:99%">
                <tr>
                    <td align="center" style="background: url(App_Themes/pwfBody/images/bottom.gif)">
						<div style="display:block; text-align:center; color:#999; margin-top:5px;">
							<p><i>&copy; M.I.S company</i> | 系統故障請打 : <font color="#FF0000">02743678060</font></p>
						</div>
					</td> 
                </tr>
            </table>
        </div> 
        
    </form>
	<script type="text/javascript">
        
		function fn_MenuHide() {
            $("#tdLeftMenu").hide();
            $("#btn_MenuShow").show();
            $("#btn_MenuHide").hide();
        }
		
        function getSelectedTabIndex() {
            var tab = document.getElementById("tabs");
            return $("#tabs").tabs('option', 'selected');

        }

        $("#ClickRefresh").click(function () {
            var index = getSelectedTabIndex();
            $fCurrent = $("#tabs").find("div iframe").eq(index);
            $fCurrent.attr("src", $fCurrent.attr("src"));
        }); 

        $("#btn_MenuShow").click(function () {
            $("#tdLeftMenu").show();
            $("#btn_MenuShow").hide();
            $("#btn_MenuHide").show();
        });

        $("#btn_MenuHide").click(function () {
            $("#tdLeftMenu").hide();
            $("#btn_MenuShow").show();
            $("#btn_MenuHide").hide();
        });
        
        $("#pnHome").click(function () {
            var mnuLink = '';
            $("#hidCurrentLink").val(mnuLink);
            $("#tabs").tabs("select", 0);
        });


        </script>
</body>
</html>
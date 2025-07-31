<%@ Page Language="C#" AutoEventWireup="true" CodeFile="LoginBody192.aspx.cs" Inherits="Pub_Module_LoginBody192" 
    Theme="pwfBody" Culture="auto" UICulture="auto"  %>
    
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
    <title></title> 
        <link rel="stylesheet" href="../../App_Themes/pwfBody/Default.css"  type="text/css" />       
	    <link rel="stylesheet" href="../Css/PccStyles.css"/>  	
        <script type="text/javascript" src="../../Pub/Js/fn_btnOver_v02.js"></script><!--滑鼠動作變換樣式-->
        <script type="text/javascript" src="../../Pub/Js/fn_imgOver.js"></script><!--滑鼠動作變換影像-->
        <script type="text/javascript" src="../../Pub/Js/fn_open.js"></script><!--視窗開啟控制-->
        <script type="text/javascript" src="../../Pub/Js/Common.js"></script>   
    </head>
    <body class="area_main">
        <form id="form1" runat="server">
            <div>
                <table width="100%">
                    <tr>
                        <td  valign="middle">
                            <asp:Label ID="title01" runat="server" CssClass="title" Text="<%$ resources: Strings, page_infor %>"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <table border="0" cellpadding="1" cellspacing="3"  width="100%">
                                <tr>
                                    <td>      
                                         <div id="Div1" class="sysMenu03_0"  runat="server">&nbsp;   
                                          <asp:HyperLink ID="hlk_Wait_Labsent" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_Labsent_title %>"></asp:HyperLink>
                                        </div> 
                                    </td>
                                </tr>
                                <tr>
                                    <td>      
                                         <div id="Div4" class="sysMenu03_0"  runat="server">&nbsp;
                                          <asp:HyperLink ID="hlk_Wait_OverTime" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_OverTime_title %>"></asp:HyperLink>
                                        </div> 
                                    </td>
                                </tr>
                                 <tr>
                                    <td>      
                                         <div id="Div5" class="sysMenu03_0"  runat="server">&nbsp;   
                                          <asp:HyperLink ID="hlk_Wait_MissCard" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_MissCard_title %>"></asp:HyperLink>
                                        </div> 
                                    </td>
                                </tr>
                                <tr>
                                    <td>      
                                         <div id="Div6" class="sysMenu03_0" runat="server">&nbsp;
                                          <asp:HyperLink ID="hlk_Wait_OverTime_Leave" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_OverTime_Leave_title %>"></asp:HyperLink>
                                        </div> 
                                    </td>
                                </tr>
                                 <tr>
			                        <td>			  
			                            <div id="Div2" class="sysMenu03_0" runat="server">&nbsp;      
                                            <asp:HyperLink ID="hlk_Wait_OverTime_App" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_OverTime_App_title %>"></asp:HyperLink>
                                           </div>
                                    </td> 
			                    </tr> 
                                <tr  >
			                        <td>			  
			                            <div id="Div3" class="sysMenu03_0" runat="server">&nbsp;      
                                            <asp:HyperLink ID="hlk_Wait_LossCard" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_Wait_LossCard_title %>"></asp:HyperLink>
                                           </div>
                                    </td> 
			                    </tr> 
                                 <tr>
                                    <td>      
                                         <div id="Div7" runat="server" class="sysMenu03_01" >&nbsp;   
                                          <asp:HyperLink ID="hlk_wait_LeaveOut" runat="server" Target="_self" Text="<%$ resources: Strings, hlk_wait_lev_out_title %>"></asp:HyperLink>
                                        </div> 
                                    </td>
                                </tr>                               
                                <tr>
                                    <td> 
                                        <div class="sysMenu03_01" >
                                            您好，PYM PLVN 安排 2019/3/29 下架，有任何建議請聯絡資訊部 陳筱婷 190-7122 shiauting@pouchen.com ，謝謝。
                                        </div>
                                        <div class="sysMenu03_01" >
                                            Xin chào, hệ thống xin nghỉ phép PLVN PYM sắp xếp ngưng dùng vào ngày 29/3/2019, có bất cứ kiến nghị gì vui lòng liên hệ Bộ Thông Tin Cô Trần Tiểu Đình 190-7122 shiauting@pouchen.com, cám ơn.
                                        </div>
                                    </td>
                                </tr>
                            </table> 
                        </td> 
                    </tr> 
                </table> 
            </div>
        </form>
    </body>
</html>

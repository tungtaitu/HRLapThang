<%@ Control Language="C#" AutoEventWireup="true" CodeFile="ddlStatus.ascx.cs" Inherits="Pub_CommControl_ddlStatus" %>
<asp:DropDownList ID="ddlStatus" runat="server" SkinID="DropDownList01">
    <asp:ListItem Value="0">---Selected---</asp:ListItem>
    <asp:ListItem Value="N">未審核(Chưa ký duyệt)</asp:ListItem>
    <asp:ListItem Value="Y">審核中(Trong thời gian ký duyệt)</asp:ListItem>
    <asp:ListItem Value="E">已核准(Đã ký duyệt)</asp:ListItem>
</asp:DropDownList>

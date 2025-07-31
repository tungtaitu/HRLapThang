<%@ Page Language="C#" AutoEventWireup="true" CodeFile="History.aspx.cs" Inherits="Pub_RemoteMethods_History" Theme="pwfBody" Culture="auto"  UICulture="auto" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:GridView ID="gv_history" runat="server" AutoGenerateColumns="False" DataKeyNames="sign_mark" OnRowDataBound="gv_history_RowDataBound" Width="100%" SkinID="GridView02" EmptyDataText="<%$ resources:Strings, not_signlog %>">
            <Columns>
                <asp:BoundField DataField="emp_name" HeaderText="<%$ resources:Strings, sign_user %>" />
                <asp:BoundField DataField="sign_date" HeaderText="<%$ resources:Strings, sign_date %>" />
                <asp:BoundField DataField="Note" HeaderText="<%$ resources:Strings, note %>">
                    <ItemStyle HorizontalAlign="Left" />
                </asp:BoundField>
                <asp:BoundField DataField="action_nm" HeaderText="<%$ resources:Strings, sign_mk %>" />
            </Columns>
            <RowStyle HorizontalAlign="Center" />
        </asp:GridView>
    
    </form>
</body>
</html>

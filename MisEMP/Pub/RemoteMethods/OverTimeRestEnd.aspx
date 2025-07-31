<%@ Page Language="C#" AutoEventWireup="true" CodeFile="OverTimeRestEnd.aspx.cs" Inherits="Pub_RemoteMethods_OverTimeRestEnd" Theme="pwfBody"  Culture="auto"  UICulture="auto" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
    <link href="../../App_Themes/pwfBody/Default.css" type="text/css" rel="stylesheet" /> 
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Literal ID="litSpliter1" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:Literal
            ID="litSpliter2" runat="server" Text="<#BreakChar#>"></asp:Literal><asp:GridView ID="grvDetail" runat="server" AutoGenerateColumns="False" OnRowDataBound="getDetail_RowDataBound" Width="100%" SkinID="GridView02" EmptyDataText="<%$ resources:Strings, not_labsent_user %>">
             <Columns>
               <asp:BoundField HeaderText="<%$ resources:Strings, number %>" />
               <asp:BoundField DataField="apply_id" HeaderText="<%$ resources:Strings, apply_id %>" />
               <asp:BoundField DataField="over_date" HeaderText="<%$ resources:Strings, over_date %>" />
               <asp:BoundField HeaderText="<%$ resources:Strings, over_date %>" />
               <asp:BoundField DataField="over_hours" HeaderText="<%$ resources:Strings, over_hours %>" >
                    <HeaderStyle Wrap="false" Width="15%" />
                    <ItemStyle Width="15%" HorizontalAlign="right" />
                 </asp:BoundField>
                 <asp:BoundField DataField="rest_hours" HeaderText="<%$ resources:Strings, rest_hours %>"  >
                    <HeaderStyle Wrap="false" Width="15%" />
                    <ItemStyle Width="15%" HorizontalAlign="right" />
                 </asp:BoundField>
                 <asp:BoundField DataField="old_rest" HeaderText="<%$ resources:Strings, old_rest %>"  >
                    <HeaderStyle Wrap="false" Width="15%" />
                    <ItemStyle Width="15%" HorizontalAlign="right" />
                 </asp:BoundField>
                 <asp:BoundField HeaderText="<%$ resources:Strings, new_rest %>"  >
                    <HeaderStyle Wrap="false" Width="15%" />
                    <ItemStyle Width="15%" HorizontalAlign="right" />
                 </asp:BoundField>
           </Columns> 
        </asp:GridView>
        </div>
    </form>
</body>
</html>
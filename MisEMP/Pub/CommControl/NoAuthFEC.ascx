<%@ Control Language="C#" AutoEventWireup="true" CodeFile="NoAuthFEC.ascx.cs" Inherits="Pub_CommControl_NoAuthFEC" %>
<asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="Conditional">
    <ContentTemplate>
        <table  width="100%">
            <tr>
                <td class="frm_itm01" width="15%"><nobr>
                    <asp:Label ID="lbl_comp" runat="server"  SkinID="itm01"></asp:Label></nobr></td>
                <td width="10%">
                    <asp:DropDownList ID="ddl_comp" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddl_SelectedIndexChanged" AppendDataBoundItems="True" SkinID="DropDownList01">
                        <asp:ListItem Value="0">---Selected---</asp:ListItem>
                    </asp:DropDownList></td>
                <td class="frm_itm01" width="10%"><nobr>
                    <asp:Label ID="lbl_fact" runat="server"  SkinID="itm01"></asp:Label></nobr></td>
                <td width="10%">
                    <asp:DropDownList ID="ddl_fact" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddl_SelectedIndexChanged" AppendDataBoundItems="True" SkinID="DropDownList01">
                        <asp:ListItem Value="0">---Selected---</asp:ListItem>
                    </asp:DropDownList></td>
                <td class="frm_itm01" width="10%"><nobr>
                    <asp:Label ID="lbl_dept" runat="server"  SkinID="itm01" ></asp:Label></nobr></td>
                <td width="10%">
                    <asp:DropDownList ID="ddl_dpt" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddl_SelectedIndexChanged" AppendDataBoundItems="True" SkinID="DropDownList01">
                        <asp:ListItem Value="0">---Selected---</asp:ListItem>
                    </asp:DropDownList></td>
                <td class="frm_itm01" width="10%"><nobr>
                    <asp:Label ID="lbl_group" runat="server"  SkinID="itm01" ></asp:Label></nobr></td>
                <td>
                    <asp:DropDownList ID="ddl_group" runat="server" AppendDataBoundItems="True" SkinID="DropDownList01" >
                        <asp:ListItem Value="0">---Selected---</asp:ListItem>
                    </asp:DropDownList></td>
            </tr>
        </table>
    </ContentTemplate>
    <Triggers>
        <asp:AsyncPostBackTrigger ControlID="ddl_comp" EventName="SelectedIndexChanged" />
        <asp:AsyncPostBackTrigger ControlID="ddl_fact" EventName="SelectedIndexChanged" />
        <asp:AsyncPostBackTrigger ControlID="ddl_group" EventName="SelectedIndexChanged" />
        <asp:AsyncPostBackTrigger ControlID="ddl_dpt" EventName="SelectedIndexChanged" />
    </Triggers>
</asp:UpdatePanel>

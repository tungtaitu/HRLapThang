<%@ Control Language="C#" AutoEventWireup="true" CodeFile="PickEmployee.ascx.cs" Inherits="Pub_BaseControls_PickEmployee" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>


<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <table>
            <tr>
                <td>
                    <asp:Button ID="btnQuery_name" runat="server" OnClick="btnQuery_name_Click" Text="查詢" />
                    <asp:Button ID="btnSelClear" runat="server" OnClick="btnSelClear_Click" Text="清空" />
                    <input id="hidFlowID" type="hidden"  runat="server"/></td>
            </tr>
            <tr>
                <td>
                    <table>
                        <tr>
                            <td align="right" bgcolor="#ffcc66" width="15%">
                                <nobr>
                                    公司別/廠別/部門/組別：</nobr></td>
                            <td bgcolor="#ffffff" colspan="3">
                                <nobr>
                                    <asp:DropDownList ID="ddlComp" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlComp_SelectedIndexChanged">
                                    </asp:DropDownList>/<asp:DropDownList ID="ddlFact" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlFact_SelectedIndexChanged">
                                    </asp:DropDownList>/<asp:DropDownList ID="ddlDept" runat="server" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlDept_SelectedIndexChanged">
                                    </asp:DropDownList><asp:DropDownList ID="ddlGroup" runat="server">
                                    </asp:DropDownList></nobr></td>
                        </tr>
                        <tr>
                            <td align="right" bgcolor="#ffcc66" width="15%">
                                姓名：</td>
                            <td bgcolor="#ffffff">
                                <asp:TextBox ID="txtQuery_name" runat="server"></asp:TextBox></td>
                            <td align="left" bgcolor="#ffffff" colspan="2">
                                <nobr>
                                    <asp:CheckBox ID="chkAbsenceUser" runat="server" Enabled="False" Text="過濾當天請假人員" /></nobr></td>
                        </tr>
                    </table>
                </td>
            </tr>
         
            <tr>
                <td align="center"  valign="top">
                    <asp:GridView ID="gv_employee" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                        BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px"
                        CellPadding="1" DataKeyNames="ID" EmptyDataText="無資料！" OnPageIndexChanging="gv_employee_PageIndexChanging"
                        Width="100%">
                        <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                        <Columns>
                            <asp:TemplateField HeaderText="勾選">
                                <ItemTemplate>
                                    <asp:CheckBox ID="chkUserID" runat="server" OnCheckedChanged="chkUserID_CheckedChanged" AutoPostBack="True" />
                                    <asp:HiddenField ID="HidID" runat="server" Value='<%# Eval("id") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:BoundField DataField="emp_no" HeaderText="員工編號" />
                            <asp:BoundField DataField="emp_name" HeaderText="員工姓名" />
                            <asp:BoundField DataField="dept_cname" HeaderText="部門名稱" />
                            <asp:BoundField DataField="group_cname" HeaderText="群組名稱" />
                        </Columns>
                        <RowStyle BackColor="White" ForeColor="#003399" />
                        <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                        <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                        <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                        <EmptyDataRowStyle BackColor="#C0FFFF" ForeColor="Red" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:HiddenField ID="hidSelEmployee" runat="server" />                    
                </td>
            </tr>
        </table>
    </ContentTemplate>
</asp:UpdatePanel>
&nbsp;&nbsp;

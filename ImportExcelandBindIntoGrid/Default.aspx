<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ImportExcelandBindIntoGrid._Default" %>



<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server" >

        <table border="1" style="width:100%";>
            <tr>
                <td>
                    <asp:Label Text="Import Excel :" ID="lblImportExcel" runat="server" Font-Bold="true"></asp:Label></td>
                <td>
                    <asp:FileUpload ID="FileUpload1" runat="server" /><asp:Button ID="btnUploadDataIntoGrid" runat="server" Text="Upload" Font-Bold="True" OnClick="btnUploadDataIntoGrid_Click" /></td>
            </tr>
            <tr>
                <td>
                    <asp:Label Text="Export Excel :" ID="lblExportExcel" runat="server" Font-Bold="true"></asp:Label></td>
                <td>
                    <asp:Button ID="btnExport" runat="server" Text="Export Excel" Font-Bold="True" OnClick="btnExport_Click" /></td>
            </tr>
            <tr><td colspan="2"> <asp:GridView ID="grdExcel" runat="server" EditRowStyle-BorderStyle="Outset" EmptyDataRowStyle-BorderStyle="Ridge" HeaderStyle-BorderStyle="Solid" AutoGenerateColumns="False" AllowSorting="True" AutoGenerateEditButton="True" CellPadding="4" ForeColor="#333333" GridLines="None" OnRowEditing="GridView1_RowEditing" OnRowCancelingEdit="GridView1_RowCancelingEdit" OnRowUpdating="GridView1_RowUpdating">
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" ></AlternatingRowStyle>
                <Columns>

                    <asp:BoundField DataField="ProductId" HeaderText="Product Id" ReadOnly="True" />
                    <%--  <asp:BoundField DataField="ProductName" HeaderText="Product Name" />--%>
                    <asp:TemplateField HeaderText="Product Name">
                        <ItemTemplate>
                            <asp:Label ID="lblProductName" runat="server" Text='<%# Bind("ProductName") %>'></asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:TextBox ID="txtProductName" runat="server" Text='<%# Bind("ProductName") %>'></asp:TextBox>
                        </EditItemTemplate>
                    </asp:TemplateField>
                    <%--  <asp:BoundField DataField="Quantity" HeaderText="Quantity" />--%>
                    <asp:TemplateField HeaderText="Product Quantity">
                        <ItemTemplate>
                            <asp:Label ID="lblQuantity" runat="server" Text='<%# Bind("Quantity") %>'></asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:TextBox ID="txtQuantity" runat="server" Text='<%# Bind("Quantity") %>'></asp:TextBox>
                        </EditItemTemplate>
                    </asp:TemplateField>
                   <%-- <asp:BoundField DataField="Price" HeaderText="Price" />--%>
                    <asp:TemplateField HeaderText="Product Price">
                        <ItemTemplate>
                            <asp:Label ID="lblPrice" runat="server" Text='<%# Bind("Price") %>'></asp:Label>
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:TextBox ID="txtPrice" runat="server" Text='<%# Bind("Price") %>'></asp:TextBox>
                        </EditItemTemplate>
                    </asp:TemplateField>


                </Columns>

                    <EditRowStyle BackColor="#999999" BorderStyle="Outset"></EditRowStyle>

                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White"></FooterStyle>

                    <HeaderStyle BackColor="#5D7B9D" BorderStyle="Solid" Font-Bold="True" ForeColor="White"></HeaderStyle>

                    <PagerStyle HorizontalAlign="Center" BackColor="#284775" ForeColor="White"></PagerStyle>

                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333"></RowStyle>

                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333"></SelectedRowStyle>

                    <SortedAscendingCellStyle BackColor="#E9E7E2"></SortedAscendingCellStyle>

                    <SortedAscendingHeaderStyle BackColor="#506C8C"></SortedAscendingHeaderStyle>

                    <SortedDescendingCellStyle BackColor="#FFFDF8"></SortedDescendingCellStyle>

                    <SortedDescendingHeaderStyle BackColor="#6F8DAE"></SortedDescendingHeaderStyle>
                </asp:GridView>
            </td></tr>
        </table>



</asp:Content>

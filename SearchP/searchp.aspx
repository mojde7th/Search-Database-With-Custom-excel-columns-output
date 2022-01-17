<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="searchp.aspx.cs" Inherits="SearchP.searchp" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <center>
<h1>جستجو</h1>
                <hr />
                <asp:TextBox ID="TextBox1" runat="server"></asp:TextBox>
                <asp:Button ID="Button1" runat="server" Text="جستجو" OnClick="Button1_Click" Width="64px" Font-Size="Large" /> 
                <asp:Button ID="Button2" runat="server" Text="دانلود اکسل" OnClick="Button2_Click" Width="98px" Font-Size="Large" />
                <asp:Button ID="Button3" runat="server" Text="دانلود اکسل با ستون‌ها" Font-Size="Large" Width="200px" OnClick="Button3_Click1"></asp:Button>
                <hr style="height: 6px; width: 1283px" />
                <asp:GridView ID="GridView1" runat="server" EmptyDataText="رکورد یافت نشد" BackColor="White" BorderColor="#CC9966" BorderStyle="None" BorderWidth="1px" CellPadding="4" Height="158px" style="margin-left: 0px" Width="98px">
                    <FooterStyle BackColor="#FFFFCC" ForeColor="#330099" />
                    <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="#FFFFCC" />
                    <PagerStyle ForeColor="#330099" HorizontalAlign="Center" BackColor="#FFFFCC" />
                    <RowStyle HorizontalAlign="Center" BackColor="White" ForeColor="#330099" />
                    <SelectedRowStyle HorizontalAlign="Center" BackColor="#FFCC66" Font-Bold="True" ForeColor="#663399" />
                    <SortedAscendingCellStyle BackColor="#FEFCEB" />
                    <SortedAscendingHeaderStyle BackColor="#AF0101" />
                    <SortedDescendingCellStyle  BackColor ="#F6F0C0" />
                    <SortedDescendingHeaderStyle BackColor="#7E0000" />
                </asp:GridView>
            </center>
        </div>
    </form>
</body>
</html>

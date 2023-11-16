<%@ Page Title="Home Page" Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="GridExcelizer._Default" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Demo GridView</title>
    <style>
        .centered-container {
            width: 60%; /* Adjust width as needed */
            margin: 0 auto; /* Center the container */
            text-align: center;
        }

        .gridview-container {
            margin-bottom: 20px; /* Space between GridView and button */
        }

        .gridview-container table {
            margin: 0 auto; /* Center-align the GridView */
            text-align: left; /* Align GridView content to the left */
        }

        /* Optional: Style for the GridView headers and cells */
        .gridview-container th, .gridview-container td {
            padding: 8px; /* Adjust padding as needed */
            border: 1px solid #ddd; /* Adjust border style as needed */
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="centered-container">
            <div class="gridview-container">
                <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false">
                    <Columns>
                        <asp:BoundField DataField="TextData" HeaderText="Text Data" />
                        <asp:BoundField DataField="NumericData" HeaderText="Numeric Data" />
                        <asp:BoundField DataField="DateData" HeaderText="Date Data" DataFormatString="{0:MM/dd/yyyy}" />
                        <asp:BoundField DataField="CurrencyData" HeaderText="Currency Data" DataFormatString="{0:C}" />
                        <asp:BoundField DataField="PercentageData" HeaderText="Percentage Data" />
                    </Columns>
                </asp:GridView>
            </div>
            <div>
                <asp:Label ID="lblMsg" runat="server" Font-Bold="True" ForeColor="Navy"></asp:Label>                      
                <asp:Button ID="ExportButton" runat="server" BackColor="#C0C0FF" Font-Bold="True"
                    Font-Size="Small" ForeColor="Black" OnClick="ExportButton_Click"
                    Text="Export To Excel" Width="110px" />
            </div>
        </div>
    </form>
</body>
</html>

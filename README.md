# GridExcelizer

## Description
GridExcelizer is a .NET Framework-based solution designed to export data from an ASP.NET GridView control into an Excel file. This project is especially useful for developers looking for a straightforward way to enable users to download GridView data in a widely-used spreadsheet format.

I created this repo as I did not find that many solutions for this when solely using OpenXML in Web Forms. If you are able to use other nuget packages please check out [ClosedXML](https://github.com/ClosedXML/ClosedXML). 

## Features
- Export GridView data to Excel format (.xlsx).
- Handle various data types including text, numbers, dates, currencies, and percentages.
- Customizable Excel formatting.
- Easy integration with ASP.NET Web Forms.

## Getting Started

### Prerequisites
- .NET Framework (Version 4.8 or later)
- ASP.NET Web Forms
- Visual Studio 

### Installation
1. Clone the repository to your local machine using `git clone <repository-url>`.
2. Open the solution in Visual Studio.
3. Build the solution to restore NuGet packages.
4. Run the application to see the demo in action.

## Usage
To use this library in your project, follow these steps:

1. **Set up the GridView**: Define your GridView in the ASP.NET page with the required columns.

   ```aspx
   <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="false">
       <!-- Define columns here -->
   </asp:GridView>

 2. **Implement Export Functionality in Code Behind**

```csharp
protected void ExportButton_Click(object sender, EventArgs e)
{
    ExcelExports exporter = new ExcelExports(Response);
    exporter.ExportToExcel(GridView1, "ExportedFileName");
}




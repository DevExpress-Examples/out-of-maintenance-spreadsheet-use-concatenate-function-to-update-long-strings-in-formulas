<!-- default badges list -->
![](https://img.shields.io/endpoint?url=https://codecentral.devexpress.com/api/v1/VersionRange/391091172/20.2.9%2B)
[![](https://img.shields.io/badge/Open_in_DevExpress_Support_Center-FF7200?style=flat-square&logo=DevExpress&logoColor=white)](https://supportcenter.devexpress.com/ticket/details/T1018328)
[![](https://img.shields.io/badge/📖_How_to_use_DevExpress_Examples-e9f6fc?style=flat-square)](https://docs.devexpress.com/GeneralInformation/403183)
<!-- default badges end -->
# Spreadsheet - How to use the CONCATENATE function to update formulas with text values longer than 255 characters

Text values in spreadsheet formulas are limited to 255 characters. If you load a document that contains formulas with string arguments longer than 255 characters, the Spreadsheet control truncates extra characters in these formulas. To avoid this behavior and allow formulas with long strings, set this property to **false**: `SpreadsheetControl.Options.Compatibility.TruncateLongStringsInFormulas`.

Consider the following restrictions if you disable this option:

* The following message appears when a user edits a formula with text values longer than 255 characters:

    ![Error Formula Message](./images/spreadsheet-message.png)

* Long strings in formulas are truncated to 255 characters when you export a workbook to Excel formats. 

We recommend that you use the CONCATENATE function or the concatenation operator (&) to create formulas with string arguments longer than 255 characters. This example demonstrates how to implement **ReplaceLongStringsWalker** that iterates through all formulas in a workbook and splits long strings in formulas with the CONCATENATE function.

<!-- default file list -->
*Files to look at*:

* [Form1.cs](./CS/SpreadsheetApp/Form1.cs) (VB: [Form1.vb](./VB/SpreadsheetApp/Form1.vb))
* [ReplaceLongStringsWalker.cs](./CS/SpreadsheetApp/ReplaceLongStringsWalker.cs) (VB: [ReplaceLongStringsWalker.vb](./VB/SpreadsheetApp/ReplaceLongStringsWalker.vb))
<!-- default file list end -->

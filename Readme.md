# Spreadsheet Mail Merge - Create a Master-Detail Report


This example demonstrates how to use the <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16257">Spreadsheet Mail Merge</a> functionality to create a master-detail template and merge data from a hierarchical data source.<br />Run the application and review the ready-to-use template (the <em>MasterDetailTemplate.xlsx</em> file) that is automatically loaded into the SpreadsheetControl.<br />Enable the <strong>Mail Merge Design View</strong> (on the <strong>Mail Merge</strong> tab) to highlight <a href="http://help.devexpress.com/#WindowsForms/CustomDocument17018">template ranges</a>.<br />- The <strong>Detail</strong> range contains fields for the top-level data from the <strong>Categories</strong> parent table of the data source. This data member is assigned via the <a href="http://help.devexpress.com/#CoreLibraries/DevExpressSpreadsheetIWorkbook_MailMergeDataMembertopic">IWorkbook.MailMergeDataMember</a> property in code.<br />- The <strong>DetailLevel0</strong> range is nested within the <strong>Detail</strong> range. It contains fields for the detail-level data from the <strong>CategoriesProducts</strong> child table. This data member is assigned via the <strong>Data Member</strong> dialog invoked when <strong>Master-Detail -> Data Member</strong> is clicked on the <strong>Mail Merge</strong> tab. <br /><br />Click the <strong>Mail Merge Preview</strong> button to display the merged document in a preview window, or click the <strong>Result</strong> button to automatically save the document to a file.


<h3>Description</h3>

You can replicate the project by performing steps of&nbsp;the following tutorial: <a href="http://help.devexpress.com/#WindowsForms/CustomDocument16965">WinForms Spreadsheet Mail Merge - Create a Master-Detail Report</a>.

<br/>



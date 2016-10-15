# New additions to Excel beta

Back in August of this year, we [announced](http://dev.office.com/blogs/power-your-apps-with-the-new-excel-rest-api) general availability of Excel REST API on Microsoft Graph. We are continually exploring opportunities to enable new scenarios and add functionality that will enable applications to build deeper and richer integration with Excel workbooks. We are happy to announce the addition of new APIs to Excel beta version. The highlights and the scenarios enabled are listed below. As with other releases, we are eager to hear your feedback. Do give this a try and let us know what you think!

The new APIs are available using the Microsoft Graph beta endpoint. You can access the new APIs for workbooks located in OneDrive for Business or in any Office 365 group as shown below:

`https://graph.microsoft.com/beta/me/drive/root:/book.xlsx:/`

## What's new

### Addition of pivotTable
Pivot table resource is now available on the `worksheet` resource. This will enable applications to refresh an individual pivot table or all pivot tables associated with a given worksheet.

* Get pivotTable: Read properties of pivotTable object. Example, `GET /worksheets/{id or name}/pivotTables/{id}`
* `refresh`: Refreshes a given pivotTable.	Example, `GET /worksheets/{id or name}/pivotTables/{id}/refresh`
* `refreshAll`: Refresh all tables within a given worksheet. Note that this action is available only on the pivot table collection. Example, 	`GET /worksheets/{id or name}/pivotTables/refreshAll`


### visibleView 
Get the range visible from a filtered range. This is available on a given range via the action: `POST /range(adress='{address}')/visibleView`


### New table resource properties

Gain insights into the structure of the Excel table with these new properties.

* `highlightFirstColumn`: Boolean; Indicates whether the first column contains special formatting.
* `highlightLastColumn`: Boolean; Indicates whether the last column contains special formatting.
* `showBandedColumns`: Boolean; Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
* `showBandedRows`: Boolean; Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
* `showFilterButton`: Boolean; Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.

### New range functions 
* `columnsAfter`: Gets a certain number of columns to the right of the given range.
* `columnsBefore`: Gets a certain number of columns to the left of the given range.
* `resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.
* `rowsAbove`: Gets a certain number of rows above a given range.
* `rowsBelow`: Gets a certain number of rows below a given range.





What to learn more?
Excellent! Visit https://dev.office.com/excel/rest, where you’ll find documentation and code samples to help you get started. It only takes a few lines of code to set up a basic integration using our to-do list sample.

Once you jump in, tell us what you think. Let us know your feedback on the API and documentation through GitHub and Stack Overflow, and make new feature suggestions on UserVoice.



Note: Any request that modifies the workbook should be performed in a persisted session. Find more details on how to create a persisted session in our documentation.

create a persisted session

POST .../workbook/CreateSessioncontent-type: Application/Json authorization: Bearer {access-token} { "persistChanges": true }



Response

HTTP code: 201, Createdcontent-type: application/json;odata.metadata

{  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.sessionInfo",  "id": "{session-id}",  "persistChanges": true}

Usage

The session ID returned from the CreateSession call is then passed as a header on subsequent API requests using the workbook-session-id HTTP header.

GET .../workbook/Worksheetsauthorization: Bearer {access-token} workbook-session-id: {session-id}

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
When a filter is applied on a table, often times it is usefult to know what values are visible (or selected from filter criteria). Until now, one had to go thorugh the underlying range and determine if the row is visible or not. This is cumbersome and hard to work when the range is fairly large. 
Hence, we've added a new API to get the visible range on a filtered range. In order to get the visible view object, simply get the object on the underlying range as below:
`GET /{range-object}/visibleView`

If the visible range happens to be large, getting the entire resource may not be performant. In such a case, iterating over the rows would provide better experience. the `rows` relationship on visibleView allows just that - iterating over the visible range rows. 

Access the rows collection using: `GET /{range-object}/visibleView/rows`

### New table resource properties

Gain insights into the structure of the Excel table with these new properties.

* `highlightFirstColumn`: Boolean; Indicates whether the first column contains special formatting.
* `highlightLastColumn`: Boolean; Indicates whether the last column contains special formatting.
* `showBandedColumns`: Boolean; Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
* `showBandedRows`: Boolean; Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
* `showFilterButton`: Boolean; Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.

### New range functions 
* `columnsAfter`: Gets a certain number of columns to the right of the given range. It takes in an optional count parameter. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `columnsBefore`: Gets a certain number of columns to the left of the given range. It takes in an optional count parameter. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns. It accepts two parameters: deltaRows: The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it. `deltaColumns`: The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
* `rowsAbove`: Gets a certain number of rows above a given range. It accepts an optional count parameter. It specifies the number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `rowsBelow`: Gets a certain number of rows below a given range. It accepts an optional count parameter. It specifies the number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.

These new range APIs can be accessed using the syntax: 
`GET /{range-object}/{function-name}`. Example, `GET /workbook/workskeets/sheet1/usedRange/columnsAfter(count=2)`  

**Note:** As always, it is a good practice to include the `Workbook-Session-Id` header if you are doing more than isolated read operations. This will ensure that the resource you may have created and modified can be accessed in the follow-up API. 
Find more details on how to create a persisted session in our documentation. In short, 

# create a persisted session
```http
POST .../workbook/CreateSession
content-type: Application/Json 
authorization: Bearer {access-token} 

{ "persistChanges": true }
```
Response

```http
HTTP code: 201, Created
content-type: application/json;odata.metadata

{  
"@odata.context": "https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.sessionInfo",  
"id": "{session-id}",  
"persistChanges": true
}
```

Usage

The session ID returned from the CreateSession call is then passed as a header on subsequent API requests using the workbook-session-id HTTP header.

```http
GET .../workbook/Worksheets
authorization: Bearer {access-token} 
workbook-session-id: {session-id}
```

#### What to learn more?
Excellent! Visit https://dev.office.com/excel/rest, where you’ll find documentation and code samples to help you get started. It only takes a few lines of code to set up a basic integration using our to-do list sample.

Once you jump in, tell us what you think. Let us know your feedback on the API and documentation through GitHub and Stack Overflow, and make new feature suggestions on UserVoice.




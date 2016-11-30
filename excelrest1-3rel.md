# New additions to the Excel REST APIs on the Microsoft Graph beta endpoint

Last August, we announced the [availability of the Excel REST API on Microsoft Graph](http://dev.office.com/blogs/power-your-apps-with-the-new-excel-rest-api). We continue to explore opportunities for new scenarios and functionality that allows for deeper and richer integration with Excel workbooks. We are happy to announce new additions to the beta Excel REST APIs. In this blog post, we'll introduce the highlights of the new APIs. As always, we are eager to hear your feedback. Try them out and let us know what you think!

The new APIs are available via the Microsoft Graph beta endpoint. You can access the new APIs for workbooks located in OneDrive for Business or in any Office 365 group.

`https://graph.microsoft.com/beta/me/drive/root:/book.xlsx:/workbook`

For an overview of the Excel REST APIs, see [Working with Excel in Microsoft Graph](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel). <!-- LG: Link to the beta topic here instead of v1.0? -->

## What's new

### Addition of pivotTable
The pivot table resource is now available on the `worksheet` resource. This will enable applications to refresh an individual pivot table or all pivot tables associated with a worksheet.

* Get pivotTable: Read properties of pivotTable object. Example, `GET /worksheets/{id or name}/pivotTables/{id}`
* `refresh`: Refreshes a given pivotTable.	Example, `GET /worksheets/{id or name}/pivotTables/{id}/refresh`
* `refreshAll`: Refresh all tables within a given worksheet. Note that this action is available only on the pivot table collection. Example, 	`GET /worksheets/{id or name}/pivotTables/refreshAll`


### visibleView 
The visibleView resource represents the visible rows of the current range. When a filter is applied on a range, it is often useful to know what values are visible (or selected from filter criteria). Previously, this required going through the underlying range to determine whether the row is visible, which is cumbersome when the range is fairly large. 
We've added a new API to get the visible range on a filtered range. To get the visible view object, get the object on the underlying range, as shown.

`GET /{range-object}/visibleView`

The visible view object <!-- LG: Clarify whether this is a resource or an object, and one word or two. -->comes with useful properties that represent the range such as cell addresses, column and row count, index, formulas, number format, values/text, and value types. 

If the visible range happens to be large, getting the entire resource might not be performant. In such a case, iterating over the rows provides a better experience. The `rows` relationship on visibleView allows just that - iterating over the visible range rows. 

Access the rows collection using the following code.

`GET /{range-object}/visibleView/rows`

### New table resource properties

Gain insights into the structure of the Excel table by using the following new properties:

* `highlightFirstColumn`: Boolean. Indicates whether the first column contains special formatting.
* `highlightLastColumn`: Boolean. Indicates whether the last column contains special formatting.
* `showBandedColumns`: Boolean. Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.
* `showBandedRows`: Boolean. Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.
* `showFilterButton`: Boolean. Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.

### New range functions 

The following new range functions have been added:

* `columnsAfter`: Gets a certain number of columns to the right of the given range. It takes in an optional `count` parameter. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `columnsBefore`: Gets a certain number of columns to the left of the given range. It takes in an optional `count` parameter. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns. It accepts two parameters: `deltaRows`: The number of rows by which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it. `deltaColumns`: The number of columnsby which to expand the bottom-right corner, relative to the current range. Use a positive number to expand the range, or a negative number to decrease it.
* `rowsAbove`: Gets a certain number of rows above a given range. It accepts an optional `count` parameter. It specifies the number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.
* `rowsBelow`: Gets a certain number of rows below a given range. It accepts an optional `count` parameter. It specifies the number of rows to include in the resulting range. In general, use a positive number to create a range outside the current range. You can also use a negative number to create a range within the current range. The default value is 1.

Access these new range APIs by using the syntax:

`GET /{range-object}/{function-name}`

For example:
`GET /workbook/workskeets/sheet1/usedRange/columnsAfter(count=2)`  

>**Note:** We recommend that you include the `Workbook-Session-Id` header if you are doing more than isolated read operations. This will ensure that the resource you created and modified can be accessed in the follow-up API. For more details about how to create a persisted session, see our documentation. <!-- LG: Link to the specific info? -->

Create a persisted session

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

#### Want to learn more?
Excellent! Visit https://dev.office.com/excel/rest, <!-- LG: Not sure about linking to this landing page, as it provides much less detail than this blog post. ;) Maybe change the heading to "Ready to get started?" --> where you’ll find documentation and code samples to help you get started. It only takes a few lines of code to set up a basic integration using our to-do list sample. <!-- LG: Link directly to the sample? -->

After you jump in, tell us what you think. Post questions or feedback about the APIs on StackOverflow, make suggestions for the docs on GitHub, or suggest new features on UserVoice.




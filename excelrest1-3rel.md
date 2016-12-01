# New additions to the Excel REST APIs on the Microsoft Graph beta endpoint

Last August, we announced the [availability of the Excel REST API on Microsoft Graph](http://dev.office.com/blogs/power-your-apps-with-the-new-excel-rest-api). We continue to explore opportunities for new scenarios and functionality that allows for deeper and richer integration with Excel workbooks. We are happy to announce new additions to the beta Excel REST APIs. In this blog post, we'll introduce the highlights of the new APIs. As always, we are eager to hear your feedback. Try them out and let us know what you think!

The new APIs are available via the Microsoft Graph beta endpoint. You can access the new APIs for workbooks located in OneDrive for Business or in any Office 365 group.

`https://graph.microsoft.com/beta/me/drive/root:/book.xlsx:/workbook`

For an overview of the Excel REST APIs, see [Working with Excel in Microsoft Graph](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel). <!-- LG: Link to the beta topic here instead of v1.0? -->

## What's new

### Pivot Table
The pivot table collection is now available on the `worksheet` resource. To begin with, this will enable developers to know about the pivot tables are available on a given workbook. This will also enable applications to refresh an individual pivot table or all pivot tables associated with a worksheet from the source data. Applications can take advantage of this feature by refreshing the pivot table remotely based on timing or other business criteria. 

* Get pivotTable: Read properties of pivotTable object. Example, `GET /worksheets/{id or name}/pivotTables/{id}`
* `refresh`: Refreshes a given pivotTable.	Example, `GET /worksheets/{id or name}/pivotTables/{id}/refresh`
* `refreshAll`: Refresh all tables within a given worksheet. Note that this action is available only on the pivot table collection. Example, 	`GET /worksheets/{id or name}/pivotTables/refreshAll`

In the future, we'll look to add more pivot table functionality using the same resources.

### visibleView 
The visibleView resource represents the visible rows of the current range. When a filter is applied on a range, it is often useful to know what values are visible (or selected from filter criteria). Previously, this required going through the underlying range to determine whether the row is visible or not. Such a procedure is cumbersome when the range is fairly large. Now, we've added a new API to get the visible range on a filtered range. 

Applications can now apply a filter on a large table based on desired criteria and get easy access to filtered data. This feature will reduce the complexity involved in determining visible rows and take advantage of Excel's filtering capability. 

In order to get the visible view object, get the object on the underlying range, as shown.

`GET /{range-object}/visibleView`

The visible view resource comes with useful properties that represent the range such as cell addresses, column and row count, index, formulas, number format, values/text, and value types. 

If the visible range happens to be large, getting the entire resource might not be performant. In such a case, iterating over the rows provides a better user experience. The `rows` relationship on visibleView allows just that - iterating over the visible range rows. 

Access the rows collection using the following code.

`GET /{range-object}/visibleView/rows`

### New range functions 

The following new range functions have been added. These functions reduce the complexity involved in determining the target range that the applications wishes to operate on.

* `columnsAfter`, `columnsBefore`, `rowsAbove`, `rowsBelow`: Gets a certain number of rows or columns positionally. 
* `resizedRange`: Gets a range object similar to the current range object, but with its bottom-right corner expanded or contracted by some number of rows and columns. 

Access these new range APIs by using the syntax:

`GET /{range-object}/{function-name}`

For example:
`GET /workbook/workskeets/sheet1/usedRange/columnsAfter(count=2)`  

### New table resource properties

Gain insights into the structure of the Excel table by using the following new properties: `highlightFirstColumn`, `highlightLastColumn`, `showBandedColumns`, `showBandedRows`, `showFilterButton`

#### Want to learn more?
Excellent! Visit https://dev.office.com/excel/rest, <!-- LG: Not sure about linking to this landing page, as it provides much less detail than this blog post. ;) Maybe change the heading to "Ready to get started?" --> where you’ll find documentation and code samples to help you get started. It only takes a few lines of code to set up a basic integration using our to-do list sample. <!-- LG: Link directly to the sample? -->

After you jump in, tell us what you think. Post questions or feedback about the APIs on StackOverflow, make suggestions for the docs on GitHub, or suggest new features on UserVoice.




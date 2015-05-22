# Set Chart Source

Returns Chart object properties. Note that the chart image is not retrned with this API.

##### HTTP Request
```
POST /Worksheets('<arg1>')/Charts('<arg2>')/setdata

```


##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------ 
| `arg1`| String | Required. Worksheet name.
| `arg2`| String | Required. Chart Name.

##### Optional Query Parameters
None

##### Optional Request Headers
None

##### Request Body
In the request body, provide a JSON object with parametrs to create a chart. 

| Parameter         | Type   |Description|
|:-----------------|:--------|:----------|
| `sourceData`  | String|  Sets an address or name of the Range object as the data source.|
| `seriesBy`  | String |  Sets the way columns or rows are used as data series on the chart. Can be one of the following e`Rows`, `Columns` or `Auto`.|


#### Example
Below operation creates a chart in Sheet1 for the given Range address and chart type. 

<!-- { "blockType": "request", "name": "set-chart-data" } -->
```
PATCH /Worksheets('Sheet1')/Charts/setdata
Content-Type: application/json
Content-Length: <length>

{
  "sourceData": "=Sheet1!$A$3:$E$6",
  "seriesBy" : "Columns"
}
```


##### Response
If successful, this method returns a [Chart](../resources/chart.md) object.

<!-- { "blockType": "response", "@odata.type": "Chart" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
	  "type": "ColumnClustered",
	  "name": "Chart1",
	  "seriesBy" : "Columns",
	  "height": 360,
	  "weight" : 216,
	  "top" : 50,
	  "left" :200
}
```


##### Error Response

Read the [Error Responses][error-response] topic for more information about how errors are returned.
[error-response]: ../misc/errors.md

 HTTP Code | HTTP Error Message | Error Code           | Error Message
:----------|:-------------------|:---------------------|:---------------------------------------------------------
 400       | Bad Request        | InvalidArgument      |The argument is invalid or missing or has an incorrect format. 
 400       | Bad Request        | InvalidRequest       | Cannot process the request.
 403       | Forbidden          | AccessDenied         | You cannot perform the requested operation.
 404       | Not Found          | ItemNotFound         | The requested resource doesn't exist.
 405       | Method Not Allowed | InvalidMethod        | The method in the request is not allowed on the resource. 
 409       | Conflict           | NameAlreadyExists    | A resource with the same name already exists.
 411       | Length Required    | ContentLengthRequired| A Content-Length header is required.
 429       |Too Many Requests        |ActivityLimitReached|Activity limit has been reached.
 500       | Internal Server Error|GeneralException    | There was an internal error while processing the request.
 503       | Service Unavailable| ServiceNotAvailable  | The service is unavailable.



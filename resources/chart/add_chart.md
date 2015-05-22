# Add Chart

Returns Chart object properties. Note that the chart image is not retrned with this API.

##### HTTP Request
```
POST /Worksheets('<arg>')/Charts/
POST /Worksheets('<arg>')/Charts/Add
```
POST /Worksheets('<arg>')/Charts/ will only add an empty chart with specified type, a SetData API will be called to relate the chart with data source.

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------ 
 `arg`| String | Required. Chart Name.

##### Optional Query Parameters
None

##### Optional Request Headers
None

##### Request Body
In the request body, provide a JSON object with parametrs to create a chart. 

| Parameter         | Type   |Description|
|:-----------------|:--------|:----------|
| `type` | String | A String value that represents the type of a chart.  |
| `sourceData`  | String | Sets an address or name of the Range object as the data source.|
| `seriesBy` | String | Sets the way columns or rows are used as data series on the chart. Can be either `Rows` or `Columns`.|

####### Optional Parameters
| Parameter         | Type   |Description|
|:-----------------|:--------|:----------|
| `name`  | String| Optional. A String value that represents a Chart object.|
| `seriesBy` | String | Optional. Returns or sets the way columns or rows are used as data series on the chart. Can be either `Rows` or `Columns`.|
| `height`|  Number |Returns integer value that represents the height, in points, of the object |
| `width` |  Number |Returns integer value that represents the width, in points, of the object. | 
| `top` |  Number |Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
| `left` |  Number |Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart). | 

#### Example
Below operation creates a chart in Sheet1 for the given Range address and chart type. 

<!-- { "blockType": "request", "name": "add-chart" } -->
```
POST /Worksheets('Sheet1')/Charts/Add
Content-Type: application/json
Content-Length: <length>
{

  "type": "ColumnClustered",
  "sourceData": "=Sheet1!$A$3:$E$6",
  "name": "Chart1",
  "seriesBy" : "Columns",
  "height": 360,
  "weight" : 216,
  "top" : 50,
  "left" :200
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
    "name": "Chart1",
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



# Update Chart

This API allows setting of Chart properties such as name, datasource, color, style, type and plotby. 

##### HTTP Request
```
PATCH /Worksheets('<arg1>')/Charts('<arg2>')
```

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------ 
 `arg1`| String | Required. Worksheet name.. 
 `arg2`| String | Required. Chart Name.

##### Optional Query Parameters
None


##### Optional Request Headers
None

##### Request Body
In the request body, provide a JSON object with parametrs to create a chart. 

| Parameter         | Type   |Description|
|:-----------------|:--------|:----------|
| `name`  | String| Optional. A String value that represents a Chart object.|
| `rangeSource`  | String | Optional. Address or name of the Range object represents the data source.|

###### Optional Parameters

| Parameter         | Type   |Description|
|:-----------------|:--------|:----------|
| `name`  | String|A String value that represents the name of a Chart object                              |
| `height`|  Number |Returns integer value that represents the height, in points, of the object |
| `width` |  Number |Returns integer value that represents the width, in points, of the object. | 
| `top` |  Number |Returns a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
| `left` |  Number |Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart). | 

#### Example
Below operation updates properties of Chart1.

<!-- { "blockType": "request", "name": "set-chart" } -->
```http
PATCH /Charts('Chart1')
Content-Type: application/json
Content-Length: <length>

{
    "name": "Chart1",
    "height": 360,
    "weight" : 216,
    "top" : 50,
    "left" :200
}
```

##### Response

If successful, this API returns a `200 OK` response to indicate that update operation was successful and the response body contains the updated view of the Chart specified.

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
 409       | Conflict           | EditConflict         | Request could not be processed because of conflict.
 411       | Length Required    | ContentLengthRequired| A Content-Length header is required.
 429       |Too Many Requests        |ActivityLimitReached|Activity limit has been reached.
 500       | Internal Server Error|GeneralException    | There was an internal error while processing the request.
 503       | Service Unavailable| ServiceNotAvailable  | The service is unavailable.


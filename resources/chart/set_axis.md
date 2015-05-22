# Update Chart Axis

This API allows setting of Chart Axis properties such as title, maximum, minimum and visibility.

##### HTTP Request
Take value axis here as an example.

```
PATCH /Worksheets('<arg1>')/Charts('<arg2>')/axes/valueaxis
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
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | Object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 
| `visible` | Boolean |True if the Axis is displayed. Read/write Boolean. | 

#### Example
Below operation updates properties of Chart1.

<!-- { "blockType": "request", "name": "set-chart-axis" } -->
```http
PATCH /Charts('Chart1')/axes/valueaxis
Content-Type: application/json
Content-Length: <length>

{
  "minimum" : 0,
  "maximum" : 100,
  "majorUnit": 5,
  "majorUnit": 1,
  "visible": true
}
```

##### Response

If successful, this API returns a `200 OK` response to indicate that update operation was successful and the response body contains the updated view of the Chart specified.

<!-- { "blockType": "response", "@odata.type": "ChartAxis" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "minimum" : 0,
  "maximum" : 100,
  "majorUnit": 5,
  "majorUnit": 1,
  "visible": true
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


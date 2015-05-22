# Get Axis Title and Related Format

Use GET method on the relevant AxisTitle object to set the Axis Title and related format.

## GET Axis Title

This API allows setting of the text of Axis Title. 

##### HTTP Request
```
GET /Charts('<arg>')/axes/valueaxis/axistitle

```

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------
 `arg`| Chart identifier | Required. Refer to `Get Chart` API for valid formats.
 

##### Optional Request Headers
None

##### Request Body

In the request body, provide a JSON object that represents the Font.

| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title of a Axis. | 
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  |

##### Example 


This example sets the text of the chart title.

<!-- { "blockType": "request", "name": "get-chart-axistitle" } -->
```http
GET /Chart('Sales')/axes/valueaxis/axistitle

```

##### Response

If successful, this method returns the [ChartAxisTitle](../../resources/chartAxisTitle.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartAxisTitle" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "text" : "Date",
  "visible" : true
}
```

##### Example 


This example sets the text of the chart title together with font and format.

<!-- { "blockType": "request", "name": "set-chart-charttitle" } -->
```http
GET /Chart('Sales')/valueaxis/axistitle

```

##### Response

If successful, this method returns the [ChartAxisTitle](../../resources/chartAxisTitle.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "AxisTitle" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
    "text" : "Date",
  "visible" : true
} 
```

##### Error Response

Read the [Error Responses][error-response] topic for more information about how errors are returned.
[error-response]: ../../misc/errors.md

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
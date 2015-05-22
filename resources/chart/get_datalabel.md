# Get Chart DataLabels and Related Format

Use GET method on the relevant ChartDataLabels object to get position and related format.

## Update Chart Legend

This API allows getting of postion and format of ChartDataLabels. 

##### HTTP Request
```
GET /Charts('<arg>')/datalabels

```

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------
 `arg`| Chart identifier | Required. Refer to `Get Chart` API for valid formats.
 

##### Optional Request Headers
None

##### Request Body
None

##### Example 


This example sets the display position of the chart DataLabels.

<!-- { "blockType": "request", "name": "get-chart-datalabels" } -->
```http
GET /Chart('Sales')/datalabels

```

##### Response

If successful, this method returns the [ChartDataLabels](../../resources/chartDataLabels.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartDataLabels" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "position" : "InsideEnd",
  "separator" : ",",
  "showBubbleSize" : false,
  "showCategoryName" : false,
  "showLegendKey" : false,
  "showPercentage" :false ,
  "showSeriesName" : true,
  "ShowValue" : true
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
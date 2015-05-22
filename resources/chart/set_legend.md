# Set Chart Legend and Related Format

Use PATCH method on the relevant Legend object to set position and related format.

## Update Chart Legend

This API allows setting of postion and format of Chart Legend. 

##### HTTP Request
```
PATCH /Charts('<arg>')/Legend

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
| `visible` | Boolean |A boolean value the represents the visibility of a ChartLegend object. If visible is set to be ture, the legend will be visible on the chart. |  
| `position` | String |Returns or sets a Legend Position value that represents the position of the legend on the chart, including `Top`,`Bottom`,`Cornor`,`Left`,`Right`,'Custom','Invalid'| 
| `overlay` | Boolean |True if the legend with be overlapping with the chart. | 

##### Example 


This example sets the display position of the chart legend.

<!-- { "blockType": "request", "name": "set-chart-legend" } -->
```http
PATCH /Chart('Sales')/Legend
Content-Type: application/json
Content-Length: <length>

{
  "visible": true,
  "position" : "Top",
  "overlay" : false
}
```

##### Response

If successful, this method returns the [ChartLegend](../../resources/chartLegend.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartLegend" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "visible": true,
  "position" : "Top",
  "overlap": False
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
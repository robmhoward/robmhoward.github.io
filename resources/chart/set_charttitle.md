# Set Chart Title and Related Format

Use PATCH method on the relevant ChartTitle object to set the Chart Title and related format.

## Update Chart Title

This API allows setting of the text of Chart Title. 

##### HTTP Request
```
PATCH /Charts('<arg>')/CharTitle

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
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. | 
| `visible` | Boolean |A boolean value the represents the visibility of a chart title object. If visible is set to be ture, the chart title will be visible on the chart. |  
| `position` | String | A constant that specifies the postition of chart title, including `Top`, `None` and `Invalid`. | 
| `overlay` | Boolean |True if the title overlays the chart. | 

##### Example 


This example sets the text of the chart title.

<!-- { "blockType": "request", "name": "set-chart-charttitle" } -->
```http
PATCH /Chart('Sales')/ChartTitle
Content-Type: application/json
Content-Length: <length>

{
  "text" : "Revenue By Quarter",
  "visible": true,
  "position" : "Top",
  "overlay" : false
}
```

##### Response

If successful, this method returns the [ChartTitle](../../resources/chartTitle.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartTitle" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "text" : "Revenue By Quarter",
  "visible": true,
  "position" : "Top",
  "overlay" : false
}
```

##### Example 


This example sets the text of the chart title together with font.

<!-- { "blockType": "request", "name": "set-chart-charttitle" } -->
```http
PATCH /Chart('Sales')/ChartTitle
Content-Type: application/json
Content-Length: <length>

{
  "text" : "Revenue By Quarter",
  "visible": true,
  "position" : "Top",
  "overlay" : false,

  "font":{          
      "bold" : true
  }
  
} 

```

##### Response

If successful, this method returns the [ChartTitle](../../resources/chartTitle.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartTitle" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "text" : "Revenue By Quarter",
  "visible": true,
  "position" : "Top",
  "overlay" : false
  
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
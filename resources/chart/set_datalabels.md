# Set Chart ChartDataLabels and Related Format

Use PATCH method on the relevant ChartDataLabels object to set position and related format.

## Update Chart Legend

This API allows setting of postion and format of ChartDataLabels. 

##### HTTP Request
```
PATCH /Charts('<arg>')/datalabels

```

##### Request Parameters
Parameter       | Type | Description
--------------- | ------ | ------------
 `arg`| Chart identifier | Required. Refer to `Get Chart` API for valid formats.
 

##### Optional Request Headers
None

##### Request Body

In the request body, provide a JSON object that represents the Font.

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`position`          |String|Returns or sets a XlDataLabelPosition value that represents the position of the data label.  |
|`separator`         |String|Sets or returns a Variant representing the separator used for the data labels on a chart. |
|`showBubbleSize`          |Boolean|True to show the bubble size for the data labels on a chart. False to hide.|
|`showCategoryName`          |Boolean|True to display the category name for the data labels on a chart. False to hide. |
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |
|`showPercentage`          |Boolean|True to display the percentage value for the data labels on a chart. False to hide.  |
|`showSeriesName`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.  |
|`ShowValue`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.|


This example sets the display position of the chart legend.

<!-- { "blockType": "request", "name": "set-chart-datalabels" } -->
```http
PATCH /Chart('Sales')/datalabels
Content-Type: application/json
Content-Length: <length>

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
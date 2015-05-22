# Set Solid Color

Set a solid fill color for fillable Chart Objects.

## Set Solid Color

This API allows getting of postion and format of lineformat. 

##### HTTP Request
```
POST /Charts('<arg>')/setsolidcolor

```

##### Request Parameters
Parameter       | Type | Description
--------------- | ------ | ------------
 `arg`| Chart identifier | Required. Refer to `Get Chart` API for valid formats.
 

##### Optional Request Headers
None

##### Request Body

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|color| String | HTML color code representing the color of the border line|

##### Example 


<!-- { "blockType": "request", "name": "set-chart-solidcolor" } -->
```http
POST /Charts("Charts1")/setsolidcolor
Content-Type: application/json
Content-Length: <length>

{
  "color" : "FF0000"
}

```

##### Response

If successful, this API returns a `204 No Content` response to indicate that
the operation was successful and there was nothing to return.

<!-- { "blockType": "response" } -->
```http
HTTP/1.1 204 No Content
````



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


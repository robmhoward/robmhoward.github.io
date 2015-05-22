# Get Chart Points and Related Format

Use Get method on the relevant ChartPoints object to get position and related format.

## Get Chart Points

This API allows getting of postion and format of an ChartPoints Object. 

##### HTTP Request
```
Get /Charts('<arg1>')/points('<arg2>')

```

##### Request Parameters
Parameter       | Type | Description
--------------- | ------ | ------------
 `arg1`| Chart identifier | Required. Refer to `Get Chart` API for valid formats.
 `arg2`| String | Required. Points Identifier
 

##### Optional Request Headers
None

##### Request Body

None

##### Example 


<!-- { "blockType": "request", "name": "get-chart-points" } -->
```http
GET /Chart('Sales')/points(1)

```

##### Response

If successful, this method returns the [ChartPoints](../../resources/chartPoints.md) object with updated values.

<!-- { "blockType": "response", "@odata.type": "ChartPoints" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>
{
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
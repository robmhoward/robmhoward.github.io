# Get Chart

Returs Chart object properties. Note that the chart image is not retrned with this API.

##### HTTP Request
```
GET /Worksheets('<arg1>')/Charts('<arg2>')
```

##### Request Parameters
|Parameter       | Type   | Description
|--------------- | ------ | ------------
| `arg1`| String | Required. Worksheet name.
| `arg2`| String | Required. Chart Name.

##### Optional Query Parameters
|Parameter       | Type   | Description|
|--------------- | ------ | ------------|
| $select| Comma-separated list of property names | |
| $expand| Comma-separated list of navigation properties to expand| Supported properties include: `title`, `series`, `axes`, `dataLabels`, `legend`, `fillFormat`,`lineFormat`,`font`|

##### Optional Request Headers
None

##### Request Body
None

#### Example
Get a Chart named Chart1

<!-- { "blockType": "request", "name": "get-chart" } -->
```http
GET /Charts('Chart1')
```

##### Response

If successful, this method returns the collection of [Chart](../../resources/chart.md) objects.

```http
HTTP/1.1 200 OK
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


##### Error Response

Read the [Error Responses][error-response] topic for more information about how errors are returned.
[error-response]: ../../misc/errors.md

 HTTP Code | HTTP Error Message | Error Code           | Error Message
:----------|:-------------------|:---------------------|:---------------------------------------------------------
 400       | Bad Request        | InvalidArgument      |The argument is invalid or missing or has an incorrect format. 
 400       | Bad Request        | InvalidRequest       | Cannot process the request.
 403       | Forbidden          | AccessDenied         | You cannot perform the requested operation.
 404       | Not Found          | ItemNotFound         | The requested resource doesn't exist.
 429       |Too Many Requests        |ActivityLimitReached|Activity limit has been reached.
 500       | Internal Server Error|GeneralException    | There was an internal error while processing the request.
 501	   | Not Implemented	| NotImplemented       | The requested feature isn't implemented.
 503       | Service Unavailable| ServiceNotAvailable  | The service is unavailable.

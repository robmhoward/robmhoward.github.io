# List Charts
Returns a collection of all the charts in the workbook.

##### HTTP Request
```
GET /Worksheet/Charts
```

##### Request Parameters
none

##### Optional Query Parameters
|Parameter       | Type   | Description|
|--------------- | ------ | ------------
| $select   | Comma-separated list of property names    | |
| $count    | Count of all Tables in the set            | |
| $expand| Comma-separated list of navigation properties to expand| Supported properties include: `ChartTitle`, `ChartArea`, `DataTable`, `Legend`, `Series`, `Axis`|


##### Optional Request Headers
None

##### Request Body
None

#### Example

Below example retrieves the charts that are part of Sheet1

<!-- { "blockType": "request", "name": "list-charts" } -->
```
GET /Worksheets("Sheet1")/Charts
```

##### Response

If successful, this method returns the collection of [Chart](../../resources/chart.md) objects.

<!-- { "blockType": "response", "@odata.type": "Chart", "isCollection": true  } -->
```http
HTTP/1.1 200 OK
Content-Type: application/json
Content-length: <length>

{
  "value": 
    [{
      "name": "Chart1",
      "height": 360,
      "weight" : 216,
      "top" : 50,
      "left" :200
    },
    {
      "name": "Chart2",
      "height": 160,
      "weight" : 216,
      "top" : 50,
      "left" :120
    },
    {
      "name": "Chart3",
      "height": 360,
      "weight" : 216,
      "top" : 150,
      "left" :220
    }]
}
```
**Note:** Response objects are truncated for clarity. All default properties 
will be returned from the actual call.

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
 501       | Not Implemented  | NotImplemented       | The requested feature isn't implemented.
 503       | Service Unavailable| ServiceNotAvailable  | The service is unavailable.


# Delete Chart

Delete a Chart from a Workbook.

##### HTTP Request
```
DELETE /Charts('<arg1>')
```

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------ 
 `arg1`| String | Required. Chart Name.

##### Optional Query Parameters
None

##### Optional Request Headers
None

##### Request Body
None

#### Example
Below operation deletes a Chart named Chart1.
<!-- { "blockType": "request", "name": "delete-chart" } -->
```http
DELETE /Chart('Chart1')
```

##### Response

If successful, this API returns a `204 No Content` response to indicate that
delete operation was successful and there was nothing to return.

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
 500       | Internal Server Error|GeneralException    | There was an internal error while processing the request.
 503       | Service Unavailable| ServiceNotAvailable  | The service is unavailable.

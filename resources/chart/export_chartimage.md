# Export Chart Image
Download chart in a graphic format. 

##### HTTP Request
```
GET /charts('<arg>''/export
```

##### Request Parameters
Parameter       | Type   | Description
--------------- | ------ | ------------ 
 `arg`| String | Required. Chart Name.


##### Optional Request Headers
None

##### Optional Request Body
In the request body, provide a JSON object with export operation's parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
filename| String | Name of the file name to be exported. Default file name is the 'name of the chart'.<extension>
format  | String | File format (JPG, GIF, TIFF, BMP, etc.). Default format is JPG. 
interactive | boolean | true to display the dialog box that contains the filter-specific options. If this argument is false, Microsoft Excel uses the default values for the filter. The default value is false.

#### Example
<!-- { "blockType": "request", "name": "delete-table" } -->
```http
GET /charts('Chart1')/export
Content-Type: application/json
Content-Length: <length>

{
  "filename": "April-Report.jpg",
  "format"  : "JPG",
  "interactive": true
}
```

##### Response

If successful, this method returns 200 OK to indicate the request has been completed and the image will attached too.

<!-- { "blockType": "response" } -->
```http
HTTP/1.1 200 OK
Content-Type: application/octet-stream
Content-Disposition: attachment; filename="April-Report.jpg"
Content-length: <length>

<raw-contents of chart imgae>
```


##### Error Response

Read the [Error Responses][error-response] topic for more information about how errors are returned.
[error-response]: ../../misc/errors.md

 HTTP Code | HTTP Error Message | Error Code           | Error Message
:----------|:-------------------|:---------------------|:---------------------------------------------------------
 400       | Bad Request        | InvalidParameter     | Supplied parameter is invalid or has incorrect format
 403       | Forbidden          | AccessRestricted     | The app does not have authorization to delete this file.
 404       | Not Found          | ResourceDoesNotExist | Resource specified in the request does not exist
 405       | Method Not Allowed | NotAllowed           | Method not allowed for the specified resource
 501       | Not Implemented    | NotImplemented       | Requested feature is not implemented

//Create a table binding to monitor data changes in the table. When data is changed, the background color of the table will be changed to orange.
 
function addEventHandler() {
 
    //Create Table1
    var ctx = new Excel.RequestContext();
    ctx.workbook.tables.add("Sheet1!A1:C4", true);
    ctx.executeAsync()
         .then(function () {
             console.log("Tablle1 Created!");
         })
         .catch(function (error) {
             console.log(JSON.stringify(error));
         });
 
    //Create a new table binding for Table1
    Office.context.document.bindings.addFromNamedItemAsync("Table1", Office.CoercionType.Table, { id: "myBinding" }, function (asyncResult) {
        if (asyncResult.status == "failed") {
            console.log("Action failed with error: " + asyncResult.error.message);
        }
        else {
            // If succeeded, then add event handler to the table binding.
            Office.select("bindings#myBinding").addHandlerAsync(Office.EventType.BindingDataChanged, onBindingDataChanged);
        }
    });
}
 
// when data in the table is changed, this event will be triggered.
function onBindingDataChanged(eventArgs) {
    var ctx = new Excel.RequestContext();
    // highlight the table in orange to indicate data has been changed.
    ctx.workbook.bindings.getItem(eventArgs.binding.id).getTable().getDataBodyRange().format.fill.color = "Orange";
    ctx.executeAsync()
        .then(function () {
            console.log("The value in this table got changed!");
        })
        .catch(function (error) {
            console.log(JSON.stringify(error));
        });
}
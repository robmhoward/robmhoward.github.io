var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Excel;
(function (Excel) {
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Application.prototype, "calculationMode", {
            get: function () {
                return this.m_calculationMode;
            },
            enumerable: true,
            configurable: true
        });

        Application.prototype.calculate = function (calculationType) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Calculate", 0 /* Default */, [calculationType]);
        };

        Application.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["CalculationMode"])) {
                this.m_calculationMode = obj["CalculationMode"];
            }
        };
        return Application;
    })(OfficeExtension.ClientObject);
    Excel.Application = Application;

    var Workbook = (function (_super) {
        __extends(Workbook, _super);
        function Workbook() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Workbook.prototype, "application", {
            get: function () {
                if (!this.m_application) {
                    this.m_application = new Excel.Application(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Application", false, false));
                }
                return this.m_application;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Workbook.prototype, "bindings", {
            get: function () {
                if (!this.m_bindings) {
                    this.m_bindings = new Excel.BindingCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Bindings", true, false));
                }
                return this.m_bindings;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Workbook.prototype, "names", {
            get: function () {
                if (!this.m_names) {
                    this.m_names = new Excel.NamedItemCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Names", false, false));
                }
                return this.m_names;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Workbook.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Workbook.prototype, "worksheets", {
            get: function () {
                if (!this.m_worksheets) {
                    this.m_worksheets = new Excel.WorksheetCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Worksheets", true, false));
                }
                return this.m_worksheets;
            },
            enumerable: true,
            configurable: true
        });

        Workbook.prototype.getSelectedRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetSelectedRange", 1 /* Read */, [], false, true));
        };

        Workbook.prototype._GetObjectByReferenceId = function (bstrReferenceId) {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_GetObjectByReferenceId", 1 /* Read */, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Workbook.prototype._GetObjectTypeNameByReferenceId = function (bstrReferenceId) {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1 /* Read */, [bstrReferenceId]);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Workbook.prototype._RemoveReference = function (bstrReferenceId) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_RemoveReference", 0 /* Default */, [bstrReferenceId]);
        };

        Workbook.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return Workbook;
    })(OfficeExtension.ClientObject);
    Excel.Workbook = Workbook;

    var Worksheet = (function (_super) {
        __extends(Worksheet, _super);
        function Worksheet() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Worksheet.prototype, "charts", {
            get: function () {
                if (!this.m_charts) {
                    this.m_charts = new Excel.ChartCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Charts", true, false));
                }
                return this.m_charts;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Worksheet.prototype, "tables", {
            get: function () {
                if (!this.m_tables) {
                    this.m_tables = new Excel.TableCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Tables", true, false));
                }
                return this.m_tables;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Worksheet.prototype, "id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Worksheet.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });

        Worksheet.prototype.activate = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Activate", 0 /* Default */, []);
        };

        Worksheet.prototype.deleteObject = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteObject", 0 /* Default */, []);
        };

        Worksheet.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetCell", 1 /* Read */, [row, column], false, true));
        };

        Worksheet.prototype.getEntireWorksheetRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetEntireWorksheetRange", 1 /* Read */, [], false, true));
        };

        Worksheet.prototype.getRange = function (address) {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [address], false, true));
        };

        Worksheet.prototype.getUsedRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetUsedRange", 1 /* Read */, [], false, true));
        };

        Worksheet.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
        };
        return Worksheet;
    })(OfficeExtension.ClientObject);
    Excel.Worksheet = Worksheet;

    var WorksheetCollection = (function (_super) {
        __extends(WorksheetCollection, _super);
        function WorksheetCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(WorksheetCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(WorksheetCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        WorksheetCollection.prototype.add = function (name) {
            return new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [name], false, true));
        };

        WorksheetCollection.prototype.getActiveWorksheet = function () {
            return new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetActiveWorksheet", 1 /* Read */, [], false, false));
        };

        WorksheetCollection.prototype.getItem = function (index) {
            return new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        WorksheetCollection.prototype.getItemAt = function (index) {
            return new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        WorksheetCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return WorksheetCollection;
    })(OfficeExtension.ClientObject);
    Excel.WorksheetCollection = WorksheetCollection;

    var Range = (function (_super) {
        __extends(Range, _super);
        function Range() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Range.prototype, "format", {
            get: function () {
                if (!this.m_format) {
                    this.m_format = new Excel.Format(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Format", false, false));
                }
                return this.m_format;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "worksheet", {
            get: function () {
                if (!this.m_worksheet) {
                    this.m_worksheet = new Excel.Worksheet(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Worksheet", false, false));
                }
                return this.m_worksheet;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "address", {
            get: function () {
                return this.m_address;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "addressLocal", {
            get: function () {
                return this.m_addressLocal;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "cellCount", {
            get: function () {
                return this.m_cellCount;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "columnCount", {
            get: function () {
                return this.m_columnCount;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "columnIndex", {
            get: function () {
                return this.m_columnIndex;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "formulas", {
            get: function () {
                return this.m_formulas;
            },
            set: function (value) {
                this.m_formulas = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Formulas", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Range.prototype, "formulasLocal", {
            get: function () {
                return this.m_formulasLocal;
            },
            set: function (value) {
                this.m_formulasLocal = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "FormulasLocal", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Range.prototype, "numberFormat", {
            get: function () {
                return this.m_numberFormat;
            },
            set: function (value) {
                this.m_numberFormat = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "NumberFormat", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Range.prototype, "rowCount", {
            get: function () {
                return this.m_rowCount;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "rowIndex", {
            get: function () {
                return this.m_rowIndex;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "text", {
            get: function () {
                return this.m_text;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "values", {
            get: function () {
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Range.prototype, "_ReferenceId", {
            get: function () {
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Range.prototype.clear = function (applyTo) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, [applyTo]);
        };

        Range.prototype.delete = function (shift) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Delete", 0 /* Default */, [shift]);
        };

        Range.prototype.getCell = function (row, column) {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetCell", 1 /* Read */, [row, column], false, true));
        };

        Range.prototype.getEntireColumn = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetEntireColumn", 1 /* Read */, [], false, true));
        };

        Range.prototype.getEntireRow = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetEntireRow", 1 /* Read */, [], false, true));
        };

        Range.prototype.getUsedRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetUsedRange", 1 /* Read */, [], false, true));
        };

        Range.prototype.insert = function (shift) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Insert", 0 /* Default */, [shift]);
        };

        Range.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        Range.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 0 /* Default */, []);
        };

        Range.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Address"])) {
                this.m_address = obj["Address"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["AddressLocal"])) {
                this.m_addressLocal = obj["AddressLocal"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["CellCount"])) {
                this.m_cellCount = obj["CellCount"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ColumnCount"])) {
                this.m_columnCount = obj["ColumnCount"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ColumnIndex"])) {
                this.m_columnIndex = obj["ColumnIndex"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Formulas"])) {
                this.m_formulas = obj["Formulas"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["FormulasLocal"])) {
                this.m_formulasLocal = obj["FormulasLocal"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["NumberFormat"])) {
                this.m_numberFormat = obj["NumberFormat"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["RowCount"])) {
                this.m_rowCount = obj["RowCount"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["RowIndex"])) {
                this.m_rowIndex = obj["RowIndex"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }
        };
        return Range;
    })(OfficeExtension.ClientObject);
    Excel.Range = Range;

    var NamedItemCollection = (function (_super) {
        __extends(NamedItemCollection, _super);
        function NamedItemCollection() {
            _super.apply(this, arguments);
        }
        NamedItemCollection.prototype.getItem = function (index) {
            return new Excel.NamedItem(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        NamedItemCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return NamedItemCollection;
    })(OfficeExtension.ClientObject);
    Excel.NamedItemCollection = NamedItemCollection;

    var NamedItem = (function (_super) {
        __extends(NamedItem, _super);
        function NamedItem() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(NamedItem.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(NamedItem.prototype, "type", {
            get: function () {
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(NamedItem.prototype, "value", {
            get: function () {
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(NamedItem.prototype, "visible", {
            get: function () {
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });


        NamedItem.prototype.getRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true));
        };

        NamedItem.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
        };
        return NamedItem;
    })(OfficeExtension.ClientObject);
    Excel.NamedItem = NamedItem;

    var Binding = (function (_super) {
        __extends(Binding, _super);
        function Binding() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Binding.prototype, "id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Binding.prototype, "type", {
            get: function () {
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });

        Binding.prototype.getRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, false));
        };

        Binding.prototype.getTable = function () {
            return new Excel.Table(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetTable", 1 /* Read */, [], false, false));
        };

        Binding.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Binding.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }
        };
        return Binding;
    })(OfficeExtension.ClientObject);
    Excel.Binding = Binding;

    var BindingCollection = (function (_super) {
        __extends(BindingCollection, _super);
        function BindingCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(BindingCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(BindingCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        BindingCollection.prototype.getItem = function (id) {
            return new Excel.Binding(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [id]));
        };

        BindingCollection.prototype.getItemAt = function (index) {
            return new Excel.Binding(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        BindingCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Binding(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return BindingCollection;
    })(OfficeExtension.ClientObject);
    Excel.BindingCollection = BindingCollection;

    var TableCollection = (function (_super) {
        __extends(TableCollection, _super);
        function TableCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        TableCollection.prototype.add = function (name, address, showHeaders, showTotals, tableStyle) {
            return new Excel.Table(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [name, address, showHeaders, showTotals, tableStyle], false, true));
        };

        TableCollection.prototype.getItem = function (id) {
            return new Excel.Table(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [id]));
        };

        TableCollection.prototype.getItemAt = function (index) {
            return new Excel.Table(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        TableCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Table(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return TableCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableCollection = TableCollection;

    var Table = (function (_super) {
        __extends(Table, _super);
        function Table() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Table.prototype, "tableColumns", {
            get: function () {
                if (!this.m_tableColumns) {
                    this.m_tableColumns = new Excel.TableColumnCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "TableColumns", true, false));
                }
                return this.m_tableColumns;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "tableRows", {
            get: function () {
                if (!this.m_tableRows) {
                    this.m_tableRows = new Excel.TableRowCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "TableRows", true, false));
                }
                return this.m_tableRows;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "showHeaders", {
            get: function () {
                return this.m_showHeaders;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "showTotals", {
            get: function () {
                return this.m_showTotals;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Table.prototype, "tableStyle", {
            get: function () {
                return this.m_tableStyle;
            },
            set: function (value) {
                this.m_tableStyle = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "TableStyle", value);
            },
            enumerable: true,
            configurable: true
        });


        Table.prototype.deleteObject = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteObject", 0 /* Default */, []);
        };

        Table.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetDataBodyRange", 1 /* Read */, [], false, true));
        };

        Table.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1 /* Read */, [], false, true));
        };

        Table.prototype.getRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true));
        };

        Table.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetTotalRowRange", 1 /* Read */, [], false, true));
        };

        Table.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowHeaders"])) {
                this.m_showHeaders = obj["ShowHeaders"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowTotals"])) {
                this.m_showTotals = obj["ShowTotals"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["TableStyle"])) {
                this.m_tableStyle = obj["TableStyle"];
            }
        };
        return Table;
    })(OfficeExtension.ClientObject);
    Excel.Table = Table;

    var TableColumnCollection = (function (_super) {
        __extends(TableColumnCollection, _super);
        function TableColumnCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableColumnCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableColumnCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        TableColumnCollection.prototype.add = function (index, values) {
            return new Excel.TableColumn(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [index, values], false, true));
        };

        TableColumnCollection.prototype.getItem = function (id) {
            return new Excel.TableColumn(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [id]));
        };

        TableColumnCollection.prototype.getItemAt = function (index) {
            return new Excel.TableColumn(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        TableColumnCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableColumn(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return TableColumnCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableColumnCollection = TableColumnCollection;

    var TableColumn = (function (_super) {
        __extends(TableColumn, _super);
        function TableColumn() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableColumn.prototype, "id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableColumn.prototype, "index", {
            get: function () {
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableColumn.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableColumn.prototype, "values", {
            get: function () {
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });


        TableColumn.prototype.deleteObject = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteObject", 0 /* Default */, []);
        };

        TableColumn.prototype.getDataBodyRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetDataBodyRange", 1 /* Read */, [], false, true));
        };

        TableColumn.prototype.getHeaderRowRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetHeaderRowRange", 1 /* Read */, [], false, true));
        };

        TableColumn.prototype.getRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true));
        };

        TableColumn.prototype.getTotalRowRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetTotalRowRange", 1 /* Read */, [], false, true));
        };

        TableColumn.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
        };
        return TableColumn;
    })(OfficeExtension.ClientObject);
    Excel.TableColumn = TableColumn;

    var TableRowCollection = (function (_super) {
        __extends(TableRowCollection, _super);
        function TableRowCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableRowCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableRowCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        TableRowCollection.prototype.add = function (index, values) {
            return new Excel.TableRow(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [index, values], false, true));
        };

        TableRowCollection.prototype.getItemAt = function (index) {
            return new Excel.TableRow(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        TableRowCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.TableRow(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return TableRowCollection;
    })(OfficeExtension.ClientObject);
    Excel.TableRowCollection = TableRowCollection;

    var TableRow = (function (_super) {
        __extends(TableRow, _super);
        function TableRow() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TableRow.prototype, "index", {
            get: function () {
                return this.m_index;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(TableRow.prototype, "values", {
            get: function () {
                return this.m_values;
            },
            set: function (value) {
                this.m_values = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Values", value);
            },
            enumerable: true,
            configurable: true
        });


        TableRow.prototype.deleteObject = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteObject", 0 /* Default */, []);
        };

        TableRow.prototype.getRange = function () {
            return new Excel.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetRange", 1 /* Read */, [], false, true));
        };

        TableRow.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Index"])) {
                this.m_index = obj["Index"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Values"])) {
                this.m_values = obj["Values"];
            }
        };
        return TableRow;
    })(OfficeExtension.ClientObject);
    Excel.TableRow = TableRow;

    var Format = (function (_super) {
        __extends(Format, _super);
        function Format() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Format.prototype, "background", {
            get: function () {
                if (!this.m_background) {
                    this.m_background = new Excel.Background(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Background", false, false));
                }
                return this.m_background;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Format.prototype, "borders", {
            get: function () {
                if (!this.m_borders) {
                    this.m_borders = new Excel.BorderCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Borders", false, false));
                }
                return this.m_borders;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Format.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.Font(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Format.prototype, "horizontalAlignment", {
            get: function () {
                return this.m_horizontalAlignment;
            },
            set: function (value) {
                this.m_horizontalAlignment = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "HorizontalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Format.prototype, "verticalAlignment", {
            get: function () {
                return this.m_verticalAlignment;
            },
            set: function (value) {
                this.m_verticalAlignment = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "VerticalAlignment", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Format.prototype, "wrapText", {
            get: function () {
                return this.m_wrapText;
            },
            set: function (value) {
                this.m_wrapText = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "WrapText", value);
            },
            enumerable: true,
            configurable: true
        });


        Format.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["HorizontalAlignment"])) {
                this.m_horizontalAlignment = obj["HorizontalAlignment"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["VerticalAlignment"])) {
                this.m_verticalAlignment = obj["VerticalAlignment"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["WrapText"])) {
                this.m_wrapText = obj["WrapText"];
            }
        };
        return Format;
    })(OfficeExtension.ClientObject);
    Excel.Format = Format;

    var Background = (function (_super) {
        __extends(Background, _super);
        function Background() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Background.prototype, "color", {
            get: function () {
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Background.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        return Background;
    })(OfficeExtension.ClientObject);
    Excel.Background = Background;

    var Border = (function (_super) {
        __extends(Border, _super);
        function Border() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Border.prototype, "color", {
            get: function () {
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Border.prototype, "lineStyle", {
            get: function () {
                return this.m_lineStyle;
            },
            set: function (value) {
                this.m_lineStyle = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LineStyle", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Border.prototype, "sideIndex", {
            get: function () {
                return this.m_sideIndex;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Border.prototype, "weight", {
            get: function () {
                return this.m_weight;
            },
            set: function (value) {
                this.m_weight = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Weight", value);
            },
            enumerable: true,
            configurable: true
        });


        Border.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LineStyle"])) {
                this.m_lineStyle = obj["LineStyle"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["SideIndex"])) {
                this.m_sideIndex = obj["SideIndex"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Weight"])) {
                this.m_weight = obj["Weight"];
            }
        };
        return Border;
    })(OfficeExtension.ClientObject);
    Excel.Border = Border;

    var BorderCollection = (function (_super) {
        __extends(BorderCollection, _super);
        function BorderCollection() {
            _super.apply(this, arguments);
        }
        BorderCollection.prototype.getItem = function (index) {
            return new Excel.Border(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        BorderCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return BorderCollection;
    })(OfficeExtension.ClientObject);
    Excel.BorderCollection = BorderCollection;

    var Font = (function (_super) {
        __extends(Font, _super);
        function Font() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Font.prototype, "bold", {
            get: function () {
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "color", {
            get: function () {
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "italic", {
            get: function () {
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "size", {
            get: function () {
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "underline", {
            get: function () {
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });


        Font.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        return Font;
    })(OfficeExtension.ClientObject);
    Excel.Font = Font;

    var ChartCollection = (function (_super) {
        __extends(ChartCollection, _super);
        function ChartCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ChartCollection.prototype.add = function (type, sourceData, seriesBy) {
            return new Excel.Chart(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [type, sourceData, seriesBy], false, true));
        };

        ChartCollection.prototype.getByName = function (Name) {
            return new Excel.Chart(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetByName", 0 /* Default */, [Name], false, false));
        };

        ChartCollection.prototype.getItem = function (id) {
            return new Excel.Chart(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [id]));
        };

        ChartCollection.prototype.getItemAt = function (index) {
            return new Excel.Chart(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        ChartCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.Chart(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return ChartCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartCollection = ChartCollection;

    var Chart = (function (_super) {
        __extends(Chart, _super);
        function Chart() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Chart.prototype, "axes", {
            get: function () {
                if (!this.m_axes) {
                    this.m_axes = new Excel.ChartAxes(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Axes", false, false));
                }
                return this.m_axes;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "dataLabels", {
            get: function () {
                if (!this.m_dataLabels) {
                    this.m_dataLabels = new Excel.ChartDataLabels(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "DataLabels", false, false));
                }
                return this.m_dataLabels;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "legend", {
            get: function () {
                if (!this.m_legend) {
                    this.m_legend = new Excel.ChartLegend(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Legend", false, false));
                }
                return this.m_legend;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "series", {
            get: function () {
                if (!this.m_series) {
                    this.m_series = new Excel.ChartSeriesCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Series", true, false));
                }
                return this.m_series;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartTitle(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "height", {
            get: function () {
                return this.m_height;
            },
            set: function (value) {
                this.m_height = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Height", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Chart.prototype, "id", {
            get: function () {
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Chart.prototype, "left", {
            get: function () {
                return this.m_left;
            },
            set: function (value) {
                this.m_left = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Left", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Chart.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Chart.prototype, "top", {
            get: function () {
                return this.m_top;
            },
            set: function (value) {
                this.m_top = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Top", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Chart.prototype, "width", {
            get: function () {
                return this.m_width;
            },
            set: function (value) {
                this.m_width = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Width", value);
            },
            enumerable: true,
            configurable: true
        });


        Chart.prototype.deleteObject = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteObject", 0 /* Default */, []);
        };

        Chart.prototype.setData = function (sourceData, seriesBy) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetData", 0 /* Default */, [sourceData, seriesBy]);
        };

        Chart.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Height"])) {
                this.m_height = obj["Height"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Left"])) {
                this.m_left = obj["Left"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Top"])) {
                this.m_top = obj["Top"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Width"])) {
                this.m_width = obj["Width"];
            }
        };
        return Chart;
    })(OfficeExtension.ClientObject);
    Excel.Chart = Chart;

    var ChartSeriesCollection = (function (_super) {
        __extends(ChartSeriesCollection, _super);
        function ChartSeriesCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartSeriesCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartSeriesCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ChartSeriesCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartSeries(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        ChartSeriesCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartSeries(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return ChartSeriesCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartSeriesCollection = ChartSeriesCollection;

    var ChartSeries = (function (_super) {
        __extends(ChartSeries, _super);
        function ChartSeries() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartSeries.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartSeries.prototype, "lineFormat", {
            get: function () {
                if (!this.m_lineFormat) {
                    this.m_lineFormat = new Excel.ChartLineFormat(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "LineFormat", false, false));
                }
                return this.m_lineFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartSeries.prototype, "points", {
            get: function () {
                if (!this.m_points) {
                    this.m_points = new Excel.ChartPointsCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Points", true, false));
                }
                return this.m_points;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartSeries.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartSeries.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
        };
        return ChartSeries;
    })(OfficeExtension.ClientObject);
    Excel.ChartSeries = ChartSeries;

    var ChartPointsCollection = (function (_super) {
        __extends(ChartPointsCollection, _super);
        function ChartPointsCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartPointsCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartPointsCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ChartPointsCollection.prototype.getItemAt = function (index) {
            return new Excel.ChartPoint(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        ChartPointsCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Excel.ChartPoint(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return ChartPointsCollection;
    })(OfficeExtension.ClientObject);
    Excel.ChartPointsCollection = ChartPointsCollection;

    var ChartPoint = (function (_super) {
        __extends(ChartPoint, _super);
        function ChartPoint() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartPoint.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        ChartPoint.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return ChartPoint;
    })(OfficeExtension.ClientObject);
    Excel.ChartPoint = ChartPoint;

    var ChartAxes = (function (_super) {
        __extends(ChartAxes, _super);
        function ChartAxes() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxes.prototype, "categoryAxis", {
            get: function () {
                if (!this.m_categoryAxis) {
                    this.m_categoryAxis = new Excel.ChartAxis(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "CategoryAxis", false, false));
                }
                return this.m_categoryAxis;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxes.prototype, "seriesAxis", {
            get: function () {
                if (!this.m_seriesAxis) {
                    this.m_seriesAxis = new Excel.ChartAxis(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "SeriesAxis", false, false));
                }
                return this.m_seriesAxis;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxes.prototype, "valueAxis", {
            get: function () {
                if (!this.m_valueAxis) {
                    this.m_valueAxis = new Excel.ChartAxis(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ValueAxis", false, false));
                }
                return this.m_valueAxis;
            },
            enumerable: true,
            configurable: true
        });

        ChartAxes.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return ChartAxes;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxes = ChartAxes;

    var ChartAxis = (function (_super) {
        __extends(ChartAxis, _super);
        function ChartAxis() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxis.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxis.prototype, "lineFormat", {
            get: function () {
                if (!this.m_lineFormat) {
                    this.m_lineFormat = new Excel.ChartLineFormat(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "LineFormat", false, false));
                }
                return this.m_lineFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxis.prototype, "majorGridlines", {
            get: function () {
                if (!this.m_majorGridlines) {
                    this.m_majorGridlines = new Excel.ChartGridlines(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "MajorGridlines", false, false));
                }
                return this.m_majorGridlines;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxis.prototype, "minorGridlines", {
            get: function () {
                if (!this.m_minorGridlines) {
                    this.m_minorGridlines = new Excel.ChartGridlines(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "MinorGridlines", false, false));
                }
                return this.m_minorGridlines;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxis.prototype, "title", {
            get: function () {
                if (!this.m_title) {
                    this.m_title = new Excel.ChartAxisTitle(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Title", false, false));
                }
                return this.m_title;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxis.prototype, "majorUnit", {
            get: function () {
                return this.m_majorUnit;
            },
            set: function (value) {
                this.m_majorUnit = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MajorUnit", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartAxis.prototype, "maximum", {
            get: function () {
                return this.m_maximum;
            },
            set: function (value) {
                this.m_maximum = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Maximum", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartAxis.prototype, "minimum", {
            get: function () {
                return this.m_minimum;
            },
            set: function (value) {
                this.m_minimum = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Minimum", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartAxis.prototype, "minorUnit", {
            get: function () {
                return this.m_minorUnit;
            },
            set: function (value) {
                this.m_minorUnit = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MinorUnit", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartAxis.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["MajorUnit"])) {
                this.m_majorUnit = obj["MajorUnit"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Maximum"])) {
                this.m_maximum = obj["Maximum"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Minimum"])) {
                this.m_minimum = obj["Minimum"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MinorUnit"])) {
                this.m_minorUnit = obj["MinorUnit"];
            }
        };
        return ChartAxis;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxis = ChartAxis;

    var ChartAxisTitle = (function (_super) {
        __extends(ChartAxisTitle, _super);
        function ChartAxisTitle() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartAxisTitle.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartAxisTitle.prototype, "text", {
            get: function () {
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartAxisTitle.prototype, "visible", {
            get: function () {
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartAxisTitle.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
        };
        return ChartAxisTitle;
    })(OfficeExtension.ClientObject);
    Excel.ChartAxisTitle = ChartAxisTitle;

    var ChartDataLabels = (function (_super) {
        __extends(ChartDataLabels, _super);
        function ChartDataLabels() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartDataLabels.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartDataLabels.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartDataLabels.prototype, "position", {
            get: function () {
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "separator", {
            get: function () {
                return this.m_separator;
            },
            set: function (value) {
                this.m_separator = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Separator", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showBubbleSize", {
            get: function () {
                return this.m_showBubbleSize;
            },
            set: function (value) {
                this.m_showBubbleSize = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowBubbleSize", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showCategoryName", {
            get: function () {
                return this.m_showCategoryName;
            },
            set: function (value) {
                this.m_showCategoryName = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowCategoryName", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showLegendKey", {
            get: function () {
                return this.m_showLegendKey;
            },
            set: function (value) {
                this.m_showLegendKey = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowLegendKey", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showPercentage", {
            get: function () {
                return this.m_showPercentage;
            },
            set: function (value) {
                this.m_showPercentage = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowPercentage", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showSeriesName", {
            get: function () {
                return this.m_showSeriesName;
            },
            set: function (value) {
                this.m_showSeriesName = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowSeriesName", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartDataLabels.prototype, "showValue", {
            get: function () {
                return this.m_showValue;
            },
            set: function (value) {
                this.m_showValue = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "ShowValue", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartDataLabels.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Separator"])) {
                this.m_separator = obj["Separator"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowBubbleSize"])) {
                this.m_showBubbleSize = obj["ShowBubbleSize"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowCategoryName"])) {
                this.m_showCategoryName = obj["ShowCategoryName"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowLegendKey"])) {
                this.m_showLegendKey = obj["ShowLegendKey"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowPercentage"])) {
                this.m_showPercentage = obj["ShowPercentage"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowSeriesName"])) {
                this.m_showSeriesName = obj["ShowSeriesName"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ShowValue"])) {
                this.m_showValue = obj["ShowValue"];
            }
        };
        return ChartDataLabels;
    })(OfficeExtension.ClientObject);
    Excel.ChartDataLabels = ChartDataLabels;

    var ChartGridlines = (function (_super) {
        __extends(ChartGridlines, _super);
        function ChartGridlines() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartGridlines.prototype, "lineFormat", {
            get: function () {
                if (!this.m_lineFormat) {
                    this.m_lineFormat = new Excel.ChartLineFormat(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "LineFormat", false, false));
                }
                return this.m_lineFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartGridlines.prototype, "visible", {
            get: function () {
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartGridlines.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
        };
        return ChartGridlines;
    })(OfficeExtension.ClientObject);
    Excel.ChartGridlines = ChartGridlines;

    var ChartLegend = (function (_super) {
        __extends(ChartLegend, _super);
        function ChartLegend() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartLegend.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartLegend.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartLegend.prototype, "overlay", {
            get: function () {
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartLegend.prototype, "position", {
            get: function () {
                return this.m_position;
            },
            set: function (value) {
                this.m_position = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Position", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartLegend.prototype, "visible", {
            get: function () {
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartLegend.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Position"])) {
                this.m_position = obj["Position"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
        };
        return ChartLegend;
    })(OfficeExtension.ClientObject);
    Excel.ChartLegend = ChartLegend;

    var ChartTitle = (function (_super) {
        __extends(ChartTitle, _super);
        function ChartTitle() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartTitle.prototype, "fillFormat", {
            get: function () {
                if (!this.m_fillFormat) {
                    this.m_fillFormat = new Excel.ChartFill(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "FillFormat", false, false));
                }
                return this.m_fillFormat;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartTitle.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Excel.ChartFont(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ChartTitle.prototype, "overlay", {
            get: function () {
                return this.m_overlay;
            },
            set: function (value) {
                this.m_overlay = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Overlay", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartTitle.prototype, "text", {
            get: function () {
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartTitle.prototype, "visible", {
            get: function () {
                return this.m_visible;
            },
            set: function (value) {
                this.m_visible = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Visible", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartTitle.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Overlay"])) {
                this.m_overlay = obj["Overlay"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Visible"])) {
                this.m_visible = obj["Visible"];
            }
        };
        return ChartTitle;
    })(OfficeExtension.ClientObject);
    Excel.ChartTitle = ChartTitle;

    var ChartFill = (function (_super) {
        __extends(ChartFill, _super);
        function ChartFill() {
            _super.apply(this, arguments);
        }
        ChartFill.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        ChartFill.prototype.setSolidColor = function (color) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetSolidColor", 0 /* Default */, [color]);
        };

        ChartFill.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
        };
        return ChartFill;
    })(OfficeExtension.ClientObject);
    Excel.ChartFill = ChartFill;

    var ChartLineFormat = (function (_super) {
        __extends(ChartLineFormat, _super);
        function ChartLineFormat() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartLineFormat.prototype, "color", {
            get: function () {
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartLineFormat.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        ChartLineFormat.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }
        };
        return ChartLineFormat;
    })(OfficeExtension.ClientObject);
    Excel.ChartLineFormat = ChartLineFormat;

    var ChartFont = (function (_super) {
        __extends(ChartFont, _super);
        function ChartFont() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartFont.prototype, "bold", {
            get: function () {
                return this.m_bold;
            },
            set: function (value) {
                this.m_bold = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Bold", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartFont.prototype, "color", {
            get: function () {
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartFont.prototype, "italic", {
            get: function () {
                return this.m_italic;
            },
            set: function (value) {
                this.m_italic = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Italic", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartFont.prototype, "name", {
            get: function () {
                return this.m_name;
            },
            set: function (value) {
                this.m_name = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Name", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartFont.prototype, "size", {
            get: function () {
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ChartFont.prototype, "underline", {
            get: function () {
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });


        ChartFont.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Bold"])) {
                this.m_bold = obj["Bold"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Italic"])) {
                this.m_italic = obj["Italic"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Size"])) {
                this.m_size = obj["Size"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }
        };
        return ChartFont;
    })(OfficeExtension.ClientObject);
    Excel.ChartFont = ChartFont;

    var BindingType = (function () {
        function BindingType() {
        }
        BindingType.range = "Range";
        BindingType.table = "Table";
        BindingType.text = "Text";
        return BindingType;
    })();
    Excel.BindingType = BindingType;

    var BorderIndex = (function () {
        function BorderIndex() {
        }
        BorderIndex.edgeTop = "EdgeTop";
        BorderIndex.edgeBottom = "EdgeBottom";
        BorderIndex.edgeLeft = "EdgeLeft";
        BorderIndex.edgeRight = "EdgeRight";
        BorderIndex.insideVertical = "InsideVertical";
        BorderIndex.insideHorizontal = "InsideHorizontal";
        BorderIndex.diagonalDown = "DiagonalDown";
        BorderIndex.diagonalUp = "DiagonalUp";
        return BorderIndex;
    })();
    Excel.BorderIndex = BorderIndex;

    var BorderLineStyle = (function () {
        function BorderLineStyle() {
        }
        BorderLineStyle.none = "None";
        BorderLineStyle.continuous = "Continuous";
        BorderLineStyle.dash = "Dash";
        BorderLineStyle.dashDot = "DashDot";
        BorderLineStyle.dashDotDot = "DashDotDot";
        BorderLineStyle.dot = "Dot";
        BorderLineStyle.double = "Double";
        BorderLineStyle.slantDashDot = "SlantDashDot";
        return BorderLineStyle;
    })();
    Excel.BorderLineStyle = BorderLineStyle;

    var BorderWeight = (function () {
        function BorderWeight() {
        }
        BorderWeight.hairline = "Hairline";
        BorderWeight.thin = "Thin";
        BorderWeight.medium = "Medium";
        BorderWeight.thick = "Thick";
        return BorderWeight;
    })();
    Excel.BorderWeight = BorderWeight;

    var CalculationMode = (function () {
        function CalculationMode() {
        }
        CalculationMode.automatic = "Automatic";
        CalculationMode.automaticExceptTables = "AutomaticExceptTables";
        CalculationMode.manual = "Manual";
        return CalculationMode;
    })();
    Excel.CalculationMode = CalculationMode;

    var CalculationType = (function () {
        function CalculationType() {
        }
        CalculationType.recalculate = "Recalculate";
        CalculationType.full = "Full";
        CalculationType.fullRebuild = "FullRebuild";
        return CalculationType;
    })();
    Excel.CalculationType = CalculationType;

    var ClearApplyTo = (function () {
        function ClearApplyTo() {
        }
        ClearApplyTo.all = "All";
        ClearApplyTo.formats = "Formats";
        ClearApplyTo.contents = "Contents";
        return ClearApplyTo;
    })();
    Excel.ClearApplyTo = ClearApplyTo;

    var ChartDataLabelPosition = (function () {
        function ChartDataLabelPosition() {
        }
        ChartDataLabelPosition.invalid = "Invalid";
        ChartDataLabelPosition.none = "None";
        ChartDataLabelPosition.center = "Center";
        ChartDataLabelPosition.insideEnd = "InsideEnd";
        ChartDataLabelPosition.insideBase = "InsideBase";
        ChartDataLabelPosition.outsideEnd = "OutsideEnd";
        ChartDataLabelPosition.left = "Left";
        ChartDataLabelPosition.right = "Right";
        ChartDataLabelPosition.top = "Top";
        ChartDataLabelPosition.bottom = "Bottom";
        ChartDataLabelPosition.bestFit = "BestFit";
        ChartDataLabelPosition.callout = "Callout";
        return ChartDataLabelPosition;
    })();
    Excel.ChartDataLabelPosition = ChartDataLabelPosition;

    var ChartLegendPosition = (function () {
        function ChartLegendPosition() {
        }
        ChartLegendPosition.invalid = "Invalid";
        ChartLegendPosition.top = "Top";
        ChartLegendPosition.bottom = "Bottom";
        ChartLegendPosition.left = "Left";
        ChartLegendPosition.right = "Right";
        ChartLegendPosition.corner = "Corner";
        ChartLegendPosition.custom = "Custom";
        return ChartLegendPosition;
    })();
    Excel.ChartLegendPosition = ChartLegendPosition;

    var ChartSeriesBy = (function () {
        function ChartSeriesBy() {
        }
        ChartSeriesBy.auto = "Auto";
        ChartSeriesBy.columns = "Columns";
        ChartSeriesBy.rows = "Rows";
        return ChartSeriesBy;
    })();
    Excel.ChartSeriesBy = ChartSeriesBy;

    var ChartType = (function () {
        function ChartType() {
        }
        ChartType.invalid = "Invalid";
        ChartType.columnClustered = "ColumnClustered";
        ChartType.columnStacked = "ColumnStacked";
        ChartType.columnStacked100 = "ColumnStacked100";
        ChartType._3DColumnClustered = "3DColumnClustered";
        ChartType._3DColumnStacked = "3DColumnStacked";
        ChartType._3DColumnStacked100 = "3DColumnStacked100";
        ChartType.barClustered = "BarClustered";
        ChartType.barStacked = "BarStacked";
        ChartType.barStacked100 = "BarStacked100";
        ChartType._3DBarClustered = "3DBarClustered";
        ChartType._3DBarStacked = "3DBarStacked";
        ChartType._3DBarStacked100 = "3DBarStacked100";
        ChartType.lineStacked = "LineStacked";
        ChartType.lineStacked100 = "LineStacked100";
        ChartType.lineMarkers = "LineMarkers";
        ChartType.lineMarkersStacked = "LineMarkersStacked";
        ChartType.lineMarkersStacked100 = "LineMarkersStacked100";
        ChartType.pieOfPie = "PieOfPie";
        ChartType.pieExploded = "PieExploded";
        ChartType._3DPieExploded = "3DPieExploded";
        ChartType.barOfPie = "BarOfPie";
        ChartType.xyscatterSmooth = "XYScatterSmooth";
        ChartType.xyscatterSmoothNoMarkers = "XYScatterSmoothNoMarkers";
        ChartType.xyscatterLines = "XYScatterLines";
        ChartType.xyscatterLinesNoMarkers = "XYScatterLinesNoMarkers";
        ChartType.areaStacked = "AreaStacked";
        ChartType.areaStacked100 = "AreaStacked100";
        ChartType._3DAreaStacked = "3DAreaStacked";
        ChartType._3DAreaStacked100 = "3DAreaStacked100";
        ChartType.doughnutExploded = "DoughnutExploded";
        ChartType.radarMarkers = "RadarMarkers";
        ChartType.radarFilled = "RadarFilled";
        ChartType.surface = "Surface";
        ChartType.surfaceWireframe = "SurfaceWireframe";
        ChartType.surfaceTopView = "SurfaceTopView";
        ChartType.surfaceTopViewWireframe = "SurfaceTopViewWireframe";
        ChartType.bubble = "Bubble";
        ChartType.bubble3DEffect = "Bubble3DEffect";
        ChartType.stockHLC = "StockHLC";
        ChartType.stockOHLC = "StockOHLC";
        ChartType.stockVHLC = "StockVHLC";
        ChartType.stockVOHLC = "StockVOHLC";
        ChartType.cylinderColClustered = "CylinderColClustered";
        ChartType.cylinderColStacked = "CylinderColStacked";
        ChartType.cylinderColStacked100 = "CylinderColStacked100";
        ChartType.cylinderBarClustered = "CylinderBarClustered";
        ChartType.cylinderBarStacked = "CylinderBarStacked";
        ChartType.cylinderBarStacked100 = "CylinderBarStacked100";
        ChartType.cylinderCol = "CylinderCol";
        ChartType.coneColClustered = "ConeColClustered";
        ChartType.coneColStacked = "ConeColStacked";
        ChartType.coneColStacked100 = "ConeColStacked100";
        ChartType.coneBarClustered = "ConeBarClustered";
        ChartType.coneBarStacked = "ConeBarStacked";
        ChartType.coneBarStacked100 = "ConeBarStacked100";
        ChartType.coneCol = "ConeCol";
        ChartType.pyramidColClustered = "PyramidColClustered";
        ChartType.pyramidColStacked = "PyramidColStacked";
        ChartType.pyramidColStacked100 = "PyramidColStacked100";
        ChartType.pyramidBarClustered = "PyramidBarClustered";
        ChartType.pyramidBarStacked = "PyramidBarStacked";
        ChartType.pyramidBarStacked100 = "PyramidBarStacked100";
        ChartType.pyramidCol = "PyramidCol";
        ChartType._3DColumn = "3DColumn";
        ChartType.line = "Line";
        ChartType._3DLine = "3DLine";
        ChartType._3DPie = "3DPie";
        ChartType.pie = "Pie";
        ChartType.xyscatter = "XYScatter";
        ChartType._3DArea = "3DArea";
        ChartType.area = "Area";
        ChartType.doughnut = "Doughnut";
        ChartType.radar = "Radar";
        return ChartType;
    })();
    Excel.ChartType = ChartType;

    var DeleteShiftDirection = (function () {
        function DeleteShiftDirection() {
        }
        DeleteShiftDirection.up = "Up";
        DeleteShiftDirection.left = "Left";
        return DeleteShiftDirection;
    })();
    Excel.DeleteShiftDirection = DeleteShiftDirection;

    var HorizontalAlignment = (function () {
        function HorizontalAlignment() {
        }
        HorizontalAlignment.general = "General";
        HorizontalAlignment.left = "Left";
        HorizontalAlignment.center = "Center";
        HorizontalAlignment.right = "Right";
        HorizontalAlignment.fill = "Fill";
        HorizontalAlignment.justify = "Justify";
        HorizontalAlignment.centerAcrossSelection = "CenterAcrossSelection";
        HorizontalAlignment.distributed = "Distributed";
        return HorizontalAlignment;
    })();
    Excel.HorizontalAlignment = HorizontalAlignment;

    var InsertShiftDirection = (function () {
        function InsertShiftDirection() {
        }
        InsertShiftDirection.down = "Down";
        InsertShiftDirection.right = "Right";
        return InsertShiftDirection;
    })();
    Excel.InsertShiftDirection = InsertShiftDirection;

    var VerticalAlignment = (function () {
        function VerticalAlignment() {
        }
        VerticalAlignment.top = "Top";
        VerticalAlignment.center = "Center";
        VerticalAlignment.bottom = "Bottom";
        VerticalAlignment.justify = "Justify";
        VerticalAlignment.distributed = "Distributed";
        return VerticalAlignment;
    })();
    Excel.VerticalAlignment = VerticalAlignment;

    var UnderlineStyle = (function () {
        function UnderlineStyle() {
        }
        UnderlineStyle.none = "None";
        UnderlineStyle.single = "Single";
        UnderlineStyle.double = "Double";
        UnderlineStyle.singleAccountant = "SingleAccountant";
        UnderlineStyle.doubleAccountant = "DoubleAccountant";
        return UnderlineStyle;
    })();
    Excel.UnderlineStyle = UnderlineStyle;

    var ErrorCodes = (function () {
        function ErrorCodes() {
        }
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.insertDeleteConflict = "InsertDeleteConflict";
        ErrorCodes.invalidArgument = "InvalidArgument";
        ErrorCodes.invalidBinding = "InvalidBinding";
        ErrorCodes.invalidOperation = "InvalidOperation";
        ErrorCodes.invalidReference = "InvalidReference";
        ErrorCodes.invalidSelection = "InvalidSelection";
        ErrorCodes.itemNotFound = "ItemNotFound";
        ErrorCodes.notImplemented = "NotImplemented";
        return ErrorCodes;
    })();
    Excel.ErrorCodes = ErrorCodes;
})(Excel || (Excel = {}));
var Excel;
(function (Excel) {
    var ExcelClientContext = (function (_super) {
        __extends(ExcelClientContext, _super);
        function ExcelClientContext(url) {
            _super.call(this, url);
            this.m_requestExecutor = new OfficeExtension.OfficeJsRequestExecutor();
            this.customData = OfficeExtension.Constants.iterativeExecutor;
            this.m_workbook = new Excel.Workbook(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
            this.rootObject = this.m_workbook;
        }
        Object.defineProperty(ExcelClientContext.prototype, "workbook", {
            get: function () {
                return this.m_workbook;
            },
            enumerable: true,
            configurable: true
        });
        return ExcelClientContext;
    })(OfficeExtension.ClientRequestContext);
    Excel.ExcelClientContext = ExcelClientContext;
})(Excel || (Excel = {}));

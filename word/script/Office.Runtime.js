var OfficeExtension;
(function (OfficeExtension) {
    var Action = (function () {
        function Action(actionInfo, isWriteOperation) {
            this.m_actionInfo = actionInfo;
            this.m_isWriteOperation = isWriteOperation;
        }
        Object.defineProperty(Action.prototype, "actionInfo", {
            get: function () {
                return this.m_actionInfo;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Action.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            enumerable: true,
            configurable: true
        });
        return Action;
    })();
    OfficeExtension.Action = Action;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ActionFactory = (function () {
        function ActionFactory() {
        }
        ActionFactory.createSetPropertyAction = function (context, parent, propertyName, value) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context.nextId(),
                ActionType: 4 /* SetProperty */,
                Name: propertyName,
                ObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var args = [value];
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var ret = new OfficeExtension.Action(actionInfo, true);
            context.pendingRequest.addAction(ret);
            context.pendingRequest.addReferencedObjectPath(parent.objectPath);
            context.pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);

            return ret;
        };

        ActionFactory.createMethodAction = function (context, parent, methodName, operationType, args) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context.nextId(),
                ActionType: 3 /* Method */,
                Name: methodName,
                ObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var referencedArgumentObjectPaths = OfficeExtension.Utility.setMethodArguments(actionInfo.ArgumentInfo, args);
            OfficeExtension.Utility.validateReferencedObjectPaths(referencedArgumentObjectPaths);
            var isWriteOperation = operationType != 1 /* Read */;
            var ret = new OfficeExtension.Action(actionInfo, isWriteOperation);
            context.pendingRequest.addAction(ret);
            context.pendingRequest.addReferencedObjectPath(parent.objectPath);
            context.pendingRequest.addReferencedObjectPaths(referencedArgumentObjectPaths);
            return ret;
        };

        ActionFactory.createQueryAction = function (context, parent, queryOption) {
            OfficeExtension.Utility.validateObjectPath(parent);
            var actionInfo = {
                Id: context.nextId(),
                ActionType: 2 /* Query */,
                Name: "",
                ObjectPathId: parent.objectPath.objectPathInfo.Id
            };
            actionInfo.QueryInfo = queryOption;
            var ret = new OfficeExtension.Action(actionInfo, false);
            context.pendingRequest.addAction(ret);
            context.pendingRequest.addReferencedObjectPath(parent.objectPath);
            return ret;
        };

        ActionFactory.createInstantiateAction = function (context, obj) {
            OfficeExtension.Utility.validateObjectPath(obj);
            var actionInfo = {
                Id: context.nextId(),
                ActionType: 1 /* Instantiate */,
                Name: "",
                ObjectPathId: obj.objectPath.objectPathInfo.Id
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context.pendingRequest.addAction(ret);
            context.pendingRequest.addReferencedObjectPath(obj.objectPath);
            context.pendingRequest.addActionResultHandler(ret, new OfficeExtension.InstantiateActionResultHandler(obj));
            return ret;
        };

        ActionFactory.createTraceAction = function (context, message) {
            var actionInfo = {
                Id: context.nextId(),
                ActionType: 5 /* Trace */,
                Name: "Trace",
                ObjectPathId: 0
            };
            var ret = new OfficeExtension.Action(actionInfo, false);
            context.pendingRequest.addAction(ret);
            context.pendingRequest.addTrace(actionInfo.Id, message);
            return ret;
        };
        return ActionFactory;
    })();
    OfficeExtension.ActionFactory = ActionFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientObject = (function () {
        function ClientObject(context, objectPath) {
            OfficeExtension.Utility.checkArgumentNull(context, "context");
            this.m_context = context;
            this.m_objectPath = objectPath;
            if (this.m_objectPath) {
                OfficeExtension.ActionFactory.createInstantiateAction(context, this);
            }
        }
        Object.defineProperty(ClientObject.prototype, "context", {
            get: function () {
                return this.m_context;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ClientObject.prototype, "objectPath", {
            get: function () {
                return this.m_objectPath;
            },
            set: function (value) {
                this.m_objectPath = value;
            },
            enumerable: true,
            configurable: true
        });


        ClientObject.prototype.handleResult = function (value) {
        };
        return ClientObject;
    })();
    OfficeExtension.ClientObject = ClientObject;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequest = (function () {
        function ClientRequest(context) {
            this.m_context = context;
            this.m_actions = [];
            this.m_actionResultHandler = {};
            this.m_referencedObjectPaths = {};
            this.m_flags = 0 /* None */;
            this.m_traceInfos = {};
        }
        Object.defineProperty(ClientRequest.prototype, "flags", {
            get: function () {
                return this.m_flags;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ClientRequest.prototype, "traceInfos", {
            get: function () {
                return this.m_traceInfos;
            },
            enumerable: true,
            configurable: true
        });

        ClientRequest.prototype.addAction = function (action) {
            if (action.isWriteOperation) {
                this.m_flags = this.m_flags | 1 /* WriteOperation */;
            }
            this.m_actions.push(action);
        };

        ClientRequest.prototype.addTrace = function (actionId, message) {
            this.m_traceInfos[actionId] = message;
        };

        ClientRequest.prototype.addReferencedObjectPath = function (objectPath) {
            if (this.m_referencedObjectPaths[objectPath.objectPathInfo.Id]) {
                return;
            }

            if (!objectPath.isValid) {
                throw Error(OfficeExtension.Utility.getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath));
            }

            while (objectPath) {
                if (objectPath.isWriteOperation) {
                    this.m_flags = this.m_flags | 1 /* WriteOperation */;
                }

                this.m_referencedObjectPaths[objectPath.objectPathInfo.Id] = objectPath;

                if (objectPath.objectPathInfo.ObjectPathType == 3 /* Method */) {
                    this.addReferencedObjectPaths(objectPath.argumentObjectPaths);
                }

                objectPath = objectPath.parentObjectPath;
            }
        };

        ClientRequest.prototype.addReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    this.addReferencedObjectPath(objectPaths[i]);
                }
            }
        };

        ClientRequest.prototype.addActionResultHandler = function (action, resultHandler) {
            this.m_actionResultHandler[action.actionInfo.Id] = resultHandler;
        };

        ClientRequest.prototype.buildRequestMessageBody = function () {
            var objectPaths = {};
            for (var i in this.m_referencedObjectPaths) {
                objectPaths[i] = this.m_referencedObjectPaths[i].objectPathInfo;
            }

            var actions = [];
            for (var index = 0; index < this.m_actions.length; index++) {
                actions.push(this.m_actions[index].actionInfo);
            }

            var ret = {
                Actions: actions,
                ObjectPaths: objectPaths
            };

            return ret;
        };

        ClientRequest.prototype.processResponse = function (msg) {
            if (msg && msg.Results) {
                for (var i = 0; i < msg.Results.length; i++) {
                    var actionResult = msg.Results[i];
                    var handler = this.m_actionResultHandler[actionResult.ActionId];
                    if (handler) {
                        handler.handleResult(actionResult.Value);
                    }
                }
            }
        };

        ClientRequest.prototype.invalidatePendingInvalidObjectPaths = function () {
            for (var i in this.m_referencedObjectPaths) {
                if (this.m_referencedObjectPaths[i].isInvalidAfterRequest) {
                    this.m_referencedObjectPaths[i].isValid = false;
                }
            }
        };
        return ClientRequest;
    })();
    OfficeExtension.ClientRequest = ClientRequest;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequestContext = (function () {
        function ClientRequestContext(url) {
            this.m_nextId = 0;
            this.m_url = url;
            if (OfficeExtension.Utility.isNullOrEmptyString(this.m_url)) {
                this.m_url = OfficeExtension.Constants.localDocument;
            }
        }
        Object.defineProperty(ClientRequestContext.prototype, "pendingRequest", {
            get: function () {
                if (this.m_pendingRequest == null) {
                    this.m_pendingRequest = new OfficeExtension.ClientRequest(this);
                }
                return this.m_pendingRequest;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ClientRequestContext.prototype, "references", {
            get: function () {
                if (!this.m_references) {
                    this.m_references = new OfficeExtension.References(this);
                }
                return this.m_references;
            },
            enumerable: true,
            configurable: true
        });

        ClientRequestContext.prototype.load = function (clientObj, option) {
            var queryOption = {};

            if (typeof (option) == "string") {
                var select = option;
                queryOption.Select = this.parseSelectExpand(select);
            } else if (typeof (option) == "object") {
                var loadOption = option;
                if (typeof (loadOption.select) == "string") {
                    queryOption.Select = this.parseSelectExpand(loadOption.select);
                }
                if (typeof (loadOption.expand) == "string") {
                    queryOption.Expand = this.parseSelectExpand(loadOption.expand);
                }
                if (typeof (loadOption.top) == "number") {
                    queryOption.Top = loadOption.top;
                }
                if (typeof (loadOption.skip) == "number") {
                    queryOption.Skip = loadOption.skip;
                }
            }

            var action = OfficeExtension.ActionFactory.createQueryAction(this, clientObj, queryOption);
            this.pendingRequest.addActionResultHandler(action, clientObj);
        };

        ClientRequestContext.prototype.trace = function (message) {
            OfficeExtension.ActionFactory.createTraceAction(this, message);
        };

        ClientRequestContext.prototype.parseSelectExpand = function (select) {
            var args = [];
            if (!OfficeExtension.Utility.isNullOrEmptyString(select)) {
                var propertyNames = select.split(",");
                for (var i = 0; i < propertyNames.length; i++) {
                    var propertyName = propertyNames[i];
                    propertyName = propertyName.trim();
                    args.push(propertyName);
                }
            }
            return args;
        };

        ClientRequestContext.prototype.executeAsyncPrivate = function (doneCallback, failCallback) {
            var req = this.pendingRequest;
            this.m_pendingRequest = null;
            var msgBody = req.buildRequestMessageBody();
            var requestFlags = req.flags;
            var requestExecutor = this.m_requestExecutor;
            if (!requestExecutor) {
                requestExecutor = new OfficeExtension.OfficeJsRequestExecutor();
            }

            var requestExecutorRequestMessage = {
                Url: this.m_url,
                Headers: null,
                Body: msgBody
            };

            req.invalidatePendingInvalidObjectPaths();

            requestExecutor.executeAsync(this.customData, requestFlags, requestExecutorRequestMessage, function (response) {
                var hasError = false;
                var result = new OfficeExtension.ClientRequestResult();
                if (!OfficeExtension.Utility.isNullOrEmptyString(response.ErrorCode)) {
                    result.errorCode = response.ErrorCode;
                    result.errorMessage = response.ErrorMessage;
                    hasError = true;
                } else if (response.Body && response.Body.Error) {
                    result.errorCode = response.Body.Error.Code;
                    result.errorMessage = response.Body.Error.Message;
                    hasError = true;
                }

                if (response.Body && response.Body.TraceIds) {
                    result.traceMessages = new Array();
                    var traceMessageMap = req.traceInfos;
                    for (var i = 0; i < response.Body.TraceIds.length; i++) {
                        var traceId = response.Body.TraceIds[i];
                        var message = traceMessageMap[traceId];
                        result.traceMessages.push(message);
                    }
                }

                if (hasError) {
                    if (failCallback) {
                        failCallback(result);
                    }
                } else {
                    req.processResponse(response.Body);
                    if (doneCallback) {
                        doneCallback(result);
                    }
                }
            });
        };

        ClientRequestContext.prototype.executeAsync = function () {
            var thisObj = this;
            var ret = new OfficeExtension.Promise(function (doneCallback, failCallback) {
                thisObj.executeAsyncPrivate(doneCallback, failCallback);
            });
            return ret;
        };

        ClientRequestContext.prototype.nextId = function () {
            return ++this.m_nextId;
        };
        return ClientRequestContext;
    })();
    OfficeExtension.ClientRequestContext = ClientRequestContext;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (ClientRequestFlags) {
        ClientRequestFlags[ClientRequestFlags["None"] = 0] = "None";
        ClientRequestFlags[ClientRequestFlags["WriteOperation"] = 1] = "WriteOperation";
    })(OfficeExtension.ClientRequestFlags || (OfficeExtension.ClientRequestFlags = {}));
    var ClientRequestFlags = OfficeExtension.ClientRequestFlags;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientRequestResult = (function () {
        function ClientRequestResult() {
        }
        return ClientRequestResult;
    })();
    OfficeExtension.ClientRequestResult = ClientRequestResult;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ClientResult = (function () {
        function ClientResult() {
        }
        Object.defineProperty(ClientResult.prototype, "value", {
            get: function () {
                return this.m_value;
            },
            enumerable: true,
            configurable: true
        });

        ClientResult.prototype.handleResult = function (value) {
            this.m_value = value;
        };
        return ClientResult;
    })();
    OfficeExtension.ClientResult = ClientResult;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Constants = (function () {
        function Constants() {
        }
        Constants.getItemAt = "GetItemAt";
        Constants.id = "Id";
        Constants.idPrivate = "_Id";
        Constants.index = "_Index";
        Constants.items = "_Items";
        Constants.iterativeExecutor = "IterativeExecutor";
        Constants.localDocument = "http://document.localhost/";
        Constants.localDocumentApiPrefix = "http://document.localhost/_api/";
        Constants.referenceId = "_ReferenceId";
        return Constants;
    })();
    OfficeExtension.Constants = Constants;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var InstantiateActionResultHandler = (function () {
        function InstantiateActionResultHandler(clientObject) {
            this.m_clientObject = clientObject;
        }
        InstantiateActionResultHandler.prototype.handleResult = function (value) {
            OfficeExtension.Utility.fixObjectPathIfNecessary(this.m_clientObject, value);
        };
        return InstantiateActionResultHandler;
    })();
    OfficeExtension.InstantiateActionResultHandler = InstantiateActionResultHandler;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (RichApiRequestMessageIndex) {
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["CustomData"] = 0] = "CustomData";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Method"] = 1] = "Method";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["PathAndQuery"] = 2] = "PathAndQuery";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Headers"] = 3] = "Headers";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["Body"] = 4] = "Body";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["AppPermission"] = 5] = "AppPermission";
        RichApiRequestMessageIndex[RichApiRequestMessageIndex["RequestFlags"] = 6] = "RequestFlags";
    })(OfficeExtension.RichApiRequestMessageIndex || (OfficeExtension.RichApiRequestMessageIndex = {}));
    var RichApiRequestMessageIndex = OfficeExtension.RichApiRequestMessageIndex;

    (function (RichApiResponseMessageIndex) {
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["StatusCode"] = 0] = "StatusCode";
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["Headers"] = 1] = "Headers";
        RichApiResponseMessageIndex[RichApiResponseMessageIndex["Body"] = 2] = "Body";
    })(OfficeExtension.RichApiResponseMessageIndex || (OfficeExtension.RichApiResponseMessageIndex = {}));
    var RichApiResponseMessageIndex = OfficeExtension.RichApiResponseMessageIndex;
    ;

    (function (ActionType) {
        ActionType[ActionType["Instantiate"] = 1] = "Instantiate";
        ActionType[ActionType["Query"] = 2] = "Query";
        ActionType[ActionType["Method"] = 3] = "Method";
        ActionType[ActionType["SetProperty"] = 4] = "SetProperty";
        ActionType[ActionType["Trace"] = 5] = "Trace";
    })(OfficeExtension.ActionType || (OfficeExtension.ActionType = {}));
    var ActionType = OfficeExtension.ActionType;

    (function (ObjectPathType) {
        ObjectPathType[ObjectPathType["GlobalObject"] = 1] = "GlobalObject";
        ObjectPathType[ObjectPathType["NewObject"] = 2] = "NewObject";
        ObjectPathType[ObjectPathType["Method"] = 3] = "Method";
        ObjectPathType[ObjectPathType["Property"] = 4] = "Property";
        ObjectPathType[ObjectPathType["Indexer"] = 5] = "Indexer";
        ObjectPathType[ObjectPathType["ReferenceId"] = 6] = "ReferenceId";
    })(OfficeExtension.ObjectPathType || (OfficeExtension.ObjectPathType = {}));
    var ObjectPathType = OfficeExtension.ObjectPathType;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPath = (function () {
        function ObjectPath(objectPathInfo, parentObjectPath, isCollection, isInvalidAfterRequest) {
            this.m_objectPathInfo = objectPathInfo;
            this.m_parentObjectPath = parentObjectPath;
            this.m_isWriteOperation = false;
            this.m_isCollection = isCollection;
            this.m_isInvalidAfterRequest = isInvalidAfterRequest;
            this.m_isValid = true;
        }
        Object.defineProperty(ObjectPath.prototype, "objectPathInfo", {
            get: function () {
                return this.m_objectPathInfo;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ObjectPath.prototype, "isWriteOperation", {
            get: function () {
                return this.m_isWriteOperation;
            },
            set: function (value) {
                this.m_isWriteOperation = value;
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ObjectPath.prototype, "isCollection", {
            get: function () {
                return this.m_isCollection;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ObjectPath.prototype, "isInvalidAfterRequest", {
            get: function () {
                return this.m_isInvalidAfterRequest;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ObjectPath.prototype, "parentObjectPath", {
            get: function () {
                return this.m_parentObjectPath;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ObjectPath.prototype, "argumentObjectPaths", {
            get: function () {
                return this.m_argumentObjectPaths;
            },
            set: function (value) {
                this.m_argumentObjectPaths = value;
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ObjectPath.prototype, "isValid", {
            get: function () {
                return this.m_isValid;
            },
            set: function (value) {
                this.m_isValid = value;
            },
            enumerable: true,
            configurable: true
        });


        ObjectPath.prototype.updateUsingObjectData = function (value) {
            var referenceId = value[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                this.m_isInvalidAfterRequest = false;
                this.m_isValid = true;
                this.m_objectPathInfo.ObjectPathType = 6 /* ReferenceId */;
                this.m_objectPathInfo.Name = referenceId;
                this.m_objectPathInfo.ArgumentInfo = {};
                this.m_parentObjectPath = null;
                this.m_argumentObjectPaths = null;
                return;
            }

            if (this.parentObjectPath && this.parentObjectPath.isCollection) {
                var id = value[OfficeExtension.Constants.id];
                if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                    id = value[OfficeExtension.Constants.idPrivate];
                }

                if (!OfficeExtension.Utility.isNullOrUndefined(id)) {
                    this.m_isInvalidAfterRequest = false;
                    this.m_isValid = true;
                    this.m_objectPathInfo.ObjectPathType = 5 /* Indexer */;
                    this.m_objectPathInfo.Name = "";
                    this.m_objectPathInfo.ArgumentInfo = {};
                    this.m_objectPathInfo.ArgumentInfo.Arguments = [id];
                    this.m_argumentObjectPaths = null;
                    return;
                }
            }
        };
        return ObjectPath;
    })();
    OfficeExtension.ObjectPath = ObjectPath;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ObjectPathFactory = (function () {
        function ObjectPathFactory() {
        }
        ObjectPathFactory.createGlobalObjectObjectPath = function (context) {
            var objectPathInfo = { Id: context.nextId(), ObjectPathType: 1 /* GlobalObject */, Name: "" };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
        };

        ObjectPathFactory.createNewObjectObjectPath = function (context, typeName, isCollection) {
            var objectPathInfo = { Id: context.nextId(), ObjectPathType: 2 /* NewObject */, Name: typeName };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, isCollection, false);
        };

        ObjectPathFactory.createPropertyObjectPath = function (context, parent, propertyName, isCollection, isInvalidAfterRequest) {
            var objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 4 /* Property */,
                Name: propertyName,
                ParentObjectPathId: parent.objectPath.objectPathInfo.Id
            };
            return new OfficeExtension.ObjectPath(objectPathInfo, parent.objectPath, isCollection, isInvalidAfterRequest);
        };

        ObjectPathFactory.createIndexerObjectPath = function (context, parent, args) {
            var objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parent.objectPath, false, false);
        };

        ObjectPathFactory.createIndexerObjectPathUsingParentPath = function (context, parentObjectPath, args) {
            var objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parentObjectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = args;
            return new OfficeExtension.ObjectPath(objectPathInfo, parentObjectPath, false, false);
        };

        ObjectPathFactory.createMethodObjectPath = function (context, parent, methodName, operationType, args, isCollection, isInvalidAfterRequest) {
            var objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 3 /* Method */,
                Name: methodName,
                ParentObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            var argumentObjectPaths = OfficeExtension.Utility.setMethodArguments(objectPathInfo.ArgumentInfo, args);
            var ret = new OfficeExtension.ObjectPath(objectPathInfo, parent.objectPath, isCollection, isInvalidAfterRequest);
            ret.argumentObjectPaths = argumentObjectPaths;
            ret.isWriteOperation = (operationType != 1 /* Read */);
            return ret;
        };

        ObjectPathFactory.createChildItemObjectPathUsingIndexer = function (context, parent, childItem) {
            var id = childItem[OfficeExtension.Constants.id];
            if (OfficeExtension.Utility.isNullOrUndefined(id)) {
                id = childItem[OfficeExtension.Constants.idPrivate];
            }

            var objectPathInfo = objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 5 /* Indexer */,
                Name: "",
                ParentObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [id];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent.objectPath, false, false);
        };

        ObjectPathFactory.createChildItemObjectPathUsingGetItemAt = function (context, parent, childItem, index) {
            var indexFromServer = childItem[OfficeExtension.Constants.index];
            if (indexFromServer) {
                index = indexFromServer;
            }

            var objectPathInfo = {
                Id: context.nextId(),
                ObjectPathType: 3 /* Method */,
                Name: OfficeExtension.Constants.getItemAt,
                ParentObjectPathId: parent.objectPath.objectPathInfo.Id,
                ArgumentInfo: {}
            };
            objectPathInfo.ArgumentInfo.Arguments = [index];
            return new OfficeExtension.ObjectPath(objectPathInfo, parent.objectPath, false, false);
        };

        ObjectPathFactory.createReferenceIdObjectPath = function (context, referenceId) {
            var objectPathInfo = { Id: context.nextId(), ObjectPathType: 6 /* ReferenceId */, Name: referenceId };
            return new OfficeExtension.ObjectPath(objectPathInfo, null, false, false);
        };
        return ObjectPathFactory;
    })();
    OfficeExtension.ObjectPathFactory = ObjectPathFactory;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var OfficeJsRequestExecutor = (function () {
        function OfficeJsRequestExecutor() {
        }
        OfficeJsRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage, callback) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            OfficeExtension.Utility.log("Request:");
            OfficeExtension.Utility.log(requestMessageText);
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", null, requestMessageText);
            OSF.DDA.RichApi.executeRichApiRequestAsync(messageSafearray, function (result) {
                OfficeExtension.Utility.log("Response:");
                OfficeExtension.Utility.log(JSON.stringify(result));
                var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
                if (result.status == "succeeded") {
                    var bodyText = OfficeExtension.RichApiMessageUtility.getResponseBody(result);
                    response.Body = JSON.parse(bodyText);
                    response.Headers = OfficeExtension.RichApiMessageUtility.getResponseHeaders(result);
                    callback(response);
                } else {
                    response.ErrorCode = result.error.code.toString();
                    response.ErrorMessage = result.error.message;
                    callback(response);
                }
            });
        };
        return OfficeJsRequestExecutor;
    })();
    OfficeExtension.OfficeJsRequestExecutor = OfficeJsRequestExecutor;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var OfficeXHRSettings = (function () {
        function OfficeXHRSettings() {
        }
        return OfficeXHRSettings;
    })();
    OfficeExtension.OfficeXHRSettings = OfficeXHRSettings;

    function resetXHRFactory(oldFactory) {
        OfficeXHR.settings.oldxhr = oldFactory;
        return officeXHRFactory;
    }
    OfficeExtension.resetXHRFactory = resetXHRFactory;

    function officeXHRFactory() {
        return new OfficeXHR;
    }
    OfficeExtension.officeXHRFactory = officeXHRFactory;

    var OfficeXHR = (function () {
        function OfficeXHR() {
        }
        OfficeXHR.prototype.open = function (method, url) {
            this.m_method = method;
            this.m_url = url;

            if (this.m_url.toLowerCase().indexOf(OfficeExtension.Constants.localDocumentApiPrefix) == 0) {
                this.m_url = this.m_url.substr(OfficeExtension.Constants.localDocumentApiPrefix.length);
            } else {
                this.m_innerXhr = OfficeXHR.settings.oldxhr();
                var thisObj = this;
                this.m_innerXhr.onreadystatechange = function () {
                    thisObj.innerXhrOnreadystatechage();
                };
                this.m_innerXhr.open(method, this.m_url);
            }
        };

        OfficeXHR.prototype.abort = function () {
            if (this.m_innerXhr) {
                this.m_innerXhr.abort();
            }
        };

        OfficeXHR.prototype.send = function (body) {
            if (this.m_innerXhr) {
                this.m_innerXhr.send(body);
            } else {
                var thisObj = this;
                var requestFlags = 0 /* None */;
                if (!OfficeExtension.Utility.isReadonlyRestRequest(this.m_method)) {
                    requestFlags = 1 /* WriteOperation */;
                }

                var execFunction = OfficeXHR.settings.executeRichApiRequestAsync;
                if (!execFunction) {
                    execFunction = OSF.DDA.RichApi.executeRichApiRequestAsync;
                }

                execFunction(OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray('', requestFlags, this.m_method, this.m_url, this.m_requestHeaders, body), function (asyncResult) {
                    thisObj.officeContextRequestCallback(asyncResult);
                });
            }
        };

        OfficeXHR.prototype.setRequestHeader = function (header, value) {
            if (this.m_innerXhr) {
                this.m_innerXhr.setRequestHeader(header, value);
            } else {
                if (!this.m_requestHeaders) {
                    this.m_requestHeaders = {};
                }

                this.m_requestHeaders[header] = value;
            }
        };

        OfficeXHR.prototype.getResponseHeader = function (header) {
            if (this.m_responseHeaders) {
                return this.m_responseHeaders[header.toUpperCase()];
            }
            return null;
        };

        OfficeXHR.prototype.getAllResponseHeaders = function () {
            return this.m_allResponseHeaders;
        };

        OfficeXHR.prototype.overrideMimeType = function (mimeType) {
            if (this.m_innerXhr) {
                this.m_innerXhr.overrideMimeType(mimeType);
            }
        };

        OfficeXHR.prototype.innerXhrOnreadystatechage = function () {
            this.readyState = this.m_innerXhr.readyState;
            if (this.readyState == OfficeXHR.DONE) {
                this.status = this.m_innerXhr.status;
                this.statusText = this.m_innerXhr.statusText;
                this.responseText = this.m_innerXhr.responseText;
                this.response = this.m_innerXhr.response;
                this.responseType = this.m_innerXhr.responseType;
                this.setAllResponseHeaders(this.m_innerXhr.getAllResponseHeaders());
            }

            if (this.onreadystatechange) {
                this.onreadystatechange();
            }
        };

        OfficeXHR.prototype.officeContextRequestCallback = function (result) {
            this.readyState = OfficeXHR.DONE;
            if (result.status == "succeeded") {
                this.status = OfficeExtension.RichApiMessageUtility.getResponseStatusCode(result);
                this.m_responseHeaders = OfficeExtension.RichApiMessageUtility.getResponseHeaders(result);
                console.debug("ResponseHeaders=" + JSON.stringify(this.m_responseHeaders));
                this.responseText = OfficeExtension.RichApiMessageUtility.getResponseBody(result);
                console.debug("ResponseText=" + this.responseText);
                this.response = this.responseText;
            } else {
                this.status = 500;
                this.statusText = "Internal Error";
            }

            if (this.onreadystatechange) {
                this.onreadystatechange();
            }
        };

        OfficeXHR.prototype.setAllResponseHeaders = function (allResponseHeaders) {
            this.m_allResponseHeaders = allResponseHeaders;
            this.m_responseHeaders = {};
            if (this.m_allResponseHeaders != null) {
                var regex = new RegExp("\r?\n");
                var entries = this.m_allResponseHeaders.split(regex);
                for (var i = 0; i < entries.length; i++) {
                    var entry = entries[i];
                    if (entry != null) {
                        var index = entry.indexOf(':');
                        if (index > 0) {
                            var key = entry.substr(0, index);
                            var value = entry.substr(index + 1);
                            key = OfficeExtension.Utility.trim(key);
                            value = OfficeExtension.Utility.trim(value);
                            this.m_responseHeaders[key.toUpperCase()] = value;
                        }
                    }
                }
            }
        };
        OfficeXHR.UNSENT = 0;
        OfficeXHR.OPENED = 1;
        OfficeXHR.DONE = 4;
        OfficeXHR.settings = new OfficeXHRSettings();
        return OfficeXHR;
    })();
    OfficeExtension.OfficeXHR = OfficeXHR;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    (function (OperationType) {
        OperationType[OperationType["Default"] = 0] = "Default";
        OperationType[OperationType["Read"] = 1] = "Read";
    })(OfficeExtension.OperationType || (OfficeExtension.OperationType = {}));
    var OperationType = OfficeExtension.OperationType;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Promise = (function () {
        function Promise(init) {
            this.m_init = init;
        }
        Promise.prototype.then = function (doneCallback, failCallback) {
            this.m_init(doneCallback, failCallback);
        };
        return Promise;
    })();
    OfficeExtension.Promise = Promise;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var References = (function () {
        function References(context) {
            this.m_context = context;
        }
        References.prototype.add = function (clientObject) {
            var referenceId = clientObject[OfficeExtension.Constants.referenceId];
            if (OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                clientObject._KeepReference();
                OfficeExtension.ActionFactory.createInstantiateAction(this.m_context, clientObject);
            }
        };

        References.prototype.remove = function (clientObject) {
            var referenceId = clientObject[OfficeExtension.Constants.referenceId];
            if (!OfficeExtension.Utility.isNullOrEmptyString(referenceId)) {
                this.m_context.rootObject._RemoveReference(referenceId);
            }
        };
        return References;
    })();
    OfficeExtension.References = References;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var ResourceStrings = (function () {
        function ResourceStrings() {
        }
        ResourceStrings.invalidObjectPath = "InvalidObjectPath";
        return ResourceStrings;
    })();
    OfficeExtension.ResourceStrings = ResourceStrings;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var RichApiMessageUtility = (function () {
        function RichApiMessageUtility() {
        }
        RichApiMessageUtility.buildRequestMessageSafeArray = function (customData, requestFlags, method, path, headers, body) {
            var headerArray = [];
            if (headers) {
                for (var headerName in headers) {
                    headerArray.push(headerName);
                    headerArray.push(headers[headerName]);
                }
            }

            var appPermission = 0;

            return [customData, method, path, headerArray, body, appPermission, requestFlags];
        };

        RichApiMessageUtility.getResponseBody = function (result) {
            return RichApiMessageUtility.getResponseBodyFromSafeArray(result.value.data);
        };

        RichApiMessageUtility.getResponseHeaders = function (result) {
            return RichApiMessageUtility.getResponseHeadersFromSafeArray(result.value.data);
        };

        RichApiMessageUtility.getResponseBodyFromSafeArray = function (data) {
            return data[2 /* Body */];
        };

        RichApiMessageUtility.getResponseHeadersFromSafeArray = function (data) {
            var arrayHeader = data[1 /* Headers */];
            if (!arrayHeader) {
                return null;
            }

            var headers = {};
            for (var i = 0; i < arrayHeader.length - 1; i += 2) {
                headers[arrayHeader[i]] = arrayHeader[i + 1];
            }

            return headers;
        };

        RichApiMessageUtility.getResponseStatusCode = function (result) {
            return RichApiMessageUtility.getResponseStatusCodeFromSafeArray(result.value.data);
        };

        RichApiMessageUtility.getResponseStatusCodeFromSafeArray = function (data) {
            return data[0 /* StatusCode */];
        };
        return RichApiMessageUtility;
    })();
    OfficeExtension.RichApiMessageUtility = RichApiMessageUtility;
})(OfficeExtension || (OfficeExtension = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var Utility = (function () {
        function Utility() {
        }
        Utility.checkArgumentNull = function (value, name) {
        };

        Utility.isNullOrUndefined = function (value) {
            if (value === null) {
                return true;
            }

            if (typeof (value) === "undefined") {
                return true;
            }

            return false;
        };

        Utility.isUndefined = function (value) {
            if (typeof (value) === "undefined") {
                return true;
            }

            return false;
        };

        Utility.isNullOrEmptyString = function (value) {
            if (value === null) {
                return true;
            }

            if (typeof (value) === "undefined") {
                return true;
            }

            if (value.length == 0) {
                return true;
            }

            return false;
        };

        Utility.trim = function (str) {
            return str.replace(new RegExp("^\\s+|\\s+$", "g"), "");
        };

        Utility.caseInsensitiveCompareString = function (str1, str2) {
            if (Utility.isNullOrUndefined(str1)) {
                return Utility.isNullOrUndefined(str2);
            } else {
                if (Utility.isNullOrUndefined(str2)) {
                    return false;
                } else {
                    return str1.toUpperCase() == str2.toUpperCase();
                }
            }
        };

        Utility.isReadonlyRestRequest = function (method) {
            return Utility.caseInsensitiveCompareString(method, "GET");
        };

        Utility.setMethodArguments = function (argumentInfo, args) {
            if (Utility.isNullOrUndefined(args)) {
                return null;
            }

            var referencedObjectPaths = new Array();
            var referencedObjectPathIds = new Array();
            var hasOne = false;
            for (var i = 0; i < args.length; i++) {
                if (args[i] instanceof OfficeExtension.ClientObject) {
                    var clientObject = args[i];
                    args[i] = clientObject.objectPath.objectPathInfo.Id;
                    referencedObjectPathIds.push(clientObject.objectPath.objectPathInfo.Id);
                    referencedObjectPaths.push(clientObject.objectPath);
                    hasOne = true;
                } else {
                    referencedObjectPathIds.push(0);
                }
            }

            argumentInfo.Arguments = args;
            if (hasOne) {
                argumentInfo.ReferencedObjectPathIds = referencedObjectPathIds;
                return referencedObjectPaths;
            }

            return null;
        };

        Utility.fixObjectPathIfNecessary = function (clientObject, value) {
            if (clientObject && clientObject.objectPath && value) {
                clientObject.objectPath.updateUsingObjectData(value);
            }
        };

        Utility.validateObjectPath = function (clientObject) {
            var objectPath = clientObject.objectPath;
            while (objectPath) {
                if (!objectPath.isValid) {
                    throw Error(Utility.getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath));
                }
                objectPath = objectPath.parentObjectPath;
            }
        };

        Utility.validateReferencedObjectPaths = function (objectPaths) {
            if (objectPaths) {
                for (var i = 0; i < objectPaths.length; i++) {
                    var objectPath = objectPaths[i];
                    while (objectPath) {
                        if (!objectPath.isValid) {
                            throw Error(Utility.getResourceString(OfficeExtension.ResourceStrings.invalidObjectPath));
                        }
                        objectPath = objectPath.parentObjectPath;
                    }
                }
            }
        };

        Utility.getResourceString = function (resourceId) {
            return resourceId;
        };

        Utility.log = function (message) {
            if (window.console && window.console.log) {
                window.console.log(message);
            }
        };
        return Utility;
    })();
    OfficeExtension.Utility = Utility;
})(OfficeExtension || (OfficeExtension = {}));

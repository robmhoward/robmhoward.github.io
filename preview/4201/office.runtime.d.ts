declare module OfficeExtension {
    class Action {
        private m_actionInfo;
        private m_isWriteOperation;
        constructor(actionInfo: ActionInfo, isWriteOperation: boolean);
        public actionInfo : ActionInfo;
        public isWriteOperation : boolean;
    }
}
declare module OfficeExtension {
    class ActionFactory {
        static createSetPropertyAction(context: ClientRequestContext, parent: ClientObject, propertyName: string, value: any): Action;
        static createMethodAction(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: any[]): Action;
        static createQueryAction(context: ClientRequestContext, parent: ClientObject, queryOption: QueryInfo): Action;
        static createInstantiateAction(context: ClientRequestContext, obj: ClientObject): Action;
        static createTraceAction(context: ClientRequestContext, message: string): Action;
    }
}
declare module OfficeExtension {
    class ClientObject implements IResultHandler {
        private m_context;
        private m_objectPath;
        constructor(context: ClientRequestContext, objectPath: ObjectPath);
        public context : ClientRequestContext;
        public objectPath : ObjectPath;
        public handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    class ClientRequest {
        private m_flags;
        private m_context;
        private m_actions;
        private m_actionResultHandler;
        private m_referencedObjectPaths;
        private m_traceInfos;
        constructor(context: ClientRequestContext);
        public flags : ClientRequestFlags;
        public traceInfos : {
            [index: number]: string;
        };
        public addAction(action: Action): void;
        public addTrace(actionId: number, message: string): void;
        public addReferencedObjectPath(objectPath: ObjectPath): void;
        public addReferencedObjectPaths(objectPaths: ObjectPath[]): void;
        public addActionResultHandler(action: Action, resultHandler: IResultHandler): void;
        public buildRequestMessageBody(): RequestMessageBody;
        public processResponse(msg: ResponseMessageBody): void;
        public invalidatePendingInvalidObjectPaths(): void;
    }
}
declare module OfficeExtension {
    interface LoadOption {
        select?: string;
        expand?: string;
        top?: number;
        skip?: number;
    }
    class ClientRequestContext {
        private m_nextId;
        private m_pendingRequest;
        private m_url;
        private m_references;
        public m_requestExecutor: IRequestExecutor;
        public rootObject: ClientObject;
        public customData: string;
        constructor(url?: string);
        public pendingRequest : ClientRequest;
        public references : References;
        public load(clientObj: ClientObject, option?: any): void;
        public trace(message: string): void;
        private parseSelectExpand(select);
        private executeAsyncPrivate(doneCallback, failCallback);
        public executeAsync(): IPromise;
        public nextId(): number;
    }
}
declare module OfficeExtension {
    enum ClientRequestFlags {
        None = 0,
        WriteOperation = 1,
    }
}
declare module OfficeExtension {
    class ClientRequestResult {
        public errorCode: string;
        public errorMessage: string;
        public traceMessages: string[];
    }
}
declare module OfficeExtension {
    class ClientResult<T> implements IResultHandler {
        private m_value;
        public value : T;
        public handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    class Constants {
        static getItemAt: string;
        static id: string;
        static idPrivate: string;
        static index: string;
        static items: string;
        static iterativeExecutor: string;
        static localDocument: string;
        static localDocumentApiPrefix: string;
        static referenceId: string;
    }
}
declare module OfficeExtension {
    class InstantiateActionResultHandler implements IResultHandler {
        private m_clientObject;
        constructor(clientObject: ClientObject);
        public handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    interface IPromise {
        then(doneCallback: (result?: any) => void, fallCallback?: (result?: any) => void): any;
    }
}
declare module OfficeExtension {
    interface IRequestExecutorRequestMessage {
        Url: string;
        Headers: {
            [name: string]: string;
        };
        Body: RequestMessageBody;
    }
    interface IRequestExecutorResponseMessage {
        ErrorCode: string;
        ErrorMessage: string;
        Headers: {
            [name: string]: string;
        };
        Body: ResponseMessageBody;
    }
    interface IRequestExecutor {
        executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage, callback: (responseMessage: IRequestExecutorResponseMessage) => void): any;
    }
}
declare module OfficeExtension {
    interface IResultHandler {
        handleResult(value: any): void;
    }
}
declare module OfficeExtension {
    enum RichApiRequestMessageIndex {
        CustomData = 0,
        Method = 1,
        PathAndQuery = 2,
        Headers = 3,
        Body = 4,
        AppPermission = 5,
        RequestFlags = 6,
    }
    enum RichApiResponseMessageIndex {
        StatusCode = 0,
        Headers = 1,
        Body = 2,
    }
    enum ActionType {
        Instantiate = 1,
        Query = 2,
        Method = 3,
        SetProperty = 4,
        Trace = 5,
    }
    enum ObjectPathType {
        GlobalObject = 1,
        NewObject = 2,
        Method = 3,
        Property = 4,
        Indexer = 5,
        ReferenceId = 6,
    }
    interface ArgumentInfo {
        Arguments?: any[];
        ReferencedObjectPathIds?: number[];
    }
    interface QueryInfo {
        Select?: string[];
        Expand?: string[];
        Skip?: number;
        Top?: number;
    }
    interface ActionInfo {
        Id: number;
        ActionType: ActionType;
        Name: string;
        ObjectPathId: number;
        ArgumentInfo?: ArgumentInfo;
        QueryInfo?: QueryInfo;
    }
    interface ActionResult {
        ActionId: number;
        Value: any;
    }
    interface ObjectPathInfo {
        Id: number;
        ObjectPathType: ObjectPathType;
        Name: string;
        ParentObjectPathId?: number;
        ArgumentInfo?: ArgumentInfo;
    }
    interface RequestMessageBody {
        Actions: ActionInfo[];
        ObjectPaths: {
            [Id: number]: ObjectPathInfo;
        };
    }
    interface ErrorInfo {
        Code: string;
        Message: string;
    }
    interface ResponseMessageBody {
        Error: ErrorInfo;
        Results: ActionResult[];
        TraceIds: number[];
    }
}
declare module OfficeExtension {
    class ObjectPath {
        private m_objectPathInfo;
        private m_isWriteOperation;
        private m_parentObjectPath;
        private m_argumentObjectPaths;
        private m_isCollection;
        private m_isInvalidAfterRequest;
        private m_isValid;
        constructor(objectPathInfo: ObjectPathInfo, parentObjectPath: ObjectPath, isCollection: boolean, isInvalidAfterRequest: boolean);
        public objectPathInfo : ObjectPathInfo;
        public isWriteOperation : boolean;
        public isCollection : boolean;
        public isInvalidAfterRequest : boolean;
        public parentObjectPath : ObjectPath;
        public argumentObjectPaths : ObjectPath[];
        public isValid : boolean;
        public updateUsingObjectData(value: Object): void;
    }
}
declare module OfficeExtension {
    class ObjectPathFactory {
        static createGlobalObjectObjectPath(context: ClientRequestContext): ObjectPath;
        static createNewObjectObjectPath(context: ClientRequestContext, typeName: string, isCollection?: boolean): ObjectPath;
        static createPropertyObjectPath(context: ClientRequestContext, parent: ClientObject, propertyName: string, isCollection?: boolean, isInvalidAfterRequest?: boolean): ObjectPath;
        static createIndexerObjectPath(context: ClientRequestContext, parent: ClientObject, args: any[]): ObjectPath;
        static createIndexerObjectPathUsingParentPath(context: ClientRequestContext, parentObjectPath: ObjectPath, args: any[]): ObjectPath;
        static createMethodObjectPath(context: ClientRequestContext, parent: ClientObject, methodName: string, operationType: OperationType, args: any[], isCollection?: boolean, isInvalidAfterRequest?: boolean): ObjectPath;
        static createChildItemObjectPathUsingIndexer(context: ClientRequestContext, parent: ClientObject, childItem: Object): ObjectPath;
        static createChildItemObjectPathUsingGetItemAt(context: ClientRequestContext, parent: ClientObject, childItem: Object, index: number): ObjectPath;
        static createReferenceIdObjectPath(context: ClientRequestContext, referenceId: string): ObjectPath;
    }
}
declare module OfficeExtension {
    class OfficeJsRequestExecutor implements IRequestExecutor {
        public executeAsync(customData: string, requestFlags: number, requestMessage: IRequestExecutorRequestMessage, callback: (responseMessage: IRequestExecutorResponseMessage) => void): void;
    }
}
declare module OfficeExtension {
    class OfficeXHRSettings {
        public oldxhr: () => IXMLHttpRequest;
        public executeRichApiRequestAsync: (message: any[], callback: (result: OSF.DDA.RichApi.ExecuteRichApiRequestResult) => void) => void;
    }
    function resetXHRFactory(oldFactory: () => IXMLHttpRequest): () => IXMLHttpRequest;
    function officeXHRFactory(): OfficeXHR;
    class OfficeXHR implements IXMLHttpRequest {
        private static UNSENT;
        private static OPENED;
        private static DONE;
        static settings: OfficeXHRSettings;
        private m_innerXhr;
        private m_method;
        private m_url;
        private m_allResponseHeaders;
        private m_responseHeaders;
        private m_requestHeaders;
        public onreadystatechange: () => void;
        public readyState: number;
        public status: number;
        public statusText: string;
        public response: any;
        public responseText: string;
        public responseType: string;
        public open(method: string, url: string): void;
        public abort(): void;
        public send(body: string): void;
        public setRequestHeader(header: string, value: string): void;
        public getResponseHeader(header: string): string;
        public getAllResponseHeaders(): string;
        public overrideMimeType(mimeType: string): void;
        private innerXhrOnreadystatechage();
        private officeContextRequestCallback(result);
        private setAllResponseHeaders(allResponseHeaders);
    }
}
declare module OfficeExtension {
    enum OperationType {
        Default = 0,
        Read = 1,
    }
}
declare module OfficeExtension {
    class Promise implements IPromise {
        private m_init;
        constructor(init: (doneCallback: (result?: any) => void, failCallback: (result?: any) => void) => void);
        public then(doneCallback: (result?: any) => void, failCallback: (result?: any) => void): void;
    }
}
declare module OfficeExtension {
    class References {
        private m_context;
        constructor(context: ClientRequestContext);
        public add(clientObject: ClientObject): void;
        public remove(clientObject: ClientObject): void;
    }
}
declare module OfficeExtension {
    class ResourceStrings {
        static invalidObjectPath: string;
    }
}
declare module OfficeExtension {
    class RichApiMessageUtility {
        static buildRequestMessageSafeArray(customData: string, requestFlags: number, method: string, path: string, headers: {
            [name: string]: string;
        }, body: string): any[];
        static getResponseBody(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): string;
        static getResponseHeaders(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): {
            [name: string]: string;
        };
        static getResponseBodyFromSafeArray(data: any[]): string;
        static getResponseHeadersFromSafeArray(data: any[]): {
            [name: string]: string;
        };
        static getResponseStatusCode(result: OSF.DDA.RichApi.ExecuteRichApiRequestResult): number;
        static getResponseStatusCodeFromSafeArray(data: any[]): number;
    }
}
declare module OfficeExtension {
    class Utility {
        static checkArgumentNull(value: any, name: string): void;
        static isNullOrUndefined(value: any): boolean;
        static isUndefined(value: any): boolean;
        static isNullOrEmptyString(value: string): boolean;
        static trim(str: string): string;
        static caseInsensitiveCompareString(str1: string, str2: string): boolean;
        static isReadonlyRestRequest(method: string): boolean;
        static setMethodArguments(argumentInfo: ArgumentInfo, args: any[]): ObjectPath[];
        static fixObjectPathIfNecessary(clientObject: ClientObject, value: Object): void;
        static validateObjectPath(clientObject: ClientObject): void;
        static validateReferencedObjectPaths(objectPaths: ObjectPath[]): void;
        static getResourceString(resourceId: string): string;
        static log(message: string): void;
    }
}

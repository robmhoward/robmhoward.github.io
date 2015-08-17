var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var Word;
(function (Word) {
    var Body = (function (_super) {
        __extends(Body, _super);
        function Body() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Body.prototype, "contentControls", {
            get: function () {
                if (!this.m_contentControls) {
                    this.m_contentControls = new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ContentControls", true, false));
                }
                return this.m_contentControls;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Body.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Word.Font(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Body.prototype, "inlinePictures", {
            get: function () {
                if (!this.m_inlinePictures) {
                    this.m_inlinePictures = new Word.InlinePictureCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
                }
                return this.m_inlinePictures;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Body.prototype, "paragraphs", {
            get: function () {
                if (!this.m_paragraphs) {
                    this.m_paragraphs = new Word.ParagraphCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
                }
                return this.m_paragraphs;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Body.prototype, "parentContentControl", {
            get: function () {
                if (!this.m_parentContentControl) {
                    this.m_parentContentControl = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
                }
                return this.m_parentContentControl;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Body.prototype, "style", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("style", this.m_style);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Body.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Body.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        Body.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Body.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Body.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Body.prototype.insertBreak = function (bt, loc) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "InsertBreak", 0 /* Default */, [bt, loc]);
        };

        Body.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, true));
        };

        Body.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, true));
        };

        Body.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, true));
        };

        Body.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, true));
        };

        Body.prototype.insertParagraph = function (paragraphText, loc) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertParagraph", 0 /* Default */, [paragraphText, loc], false, true));
        };

        Body.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, true));
        };

        Body.prototype.search = function (searchText, searchOptions) {
            return new Word.SearchResultCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Search", 1 /* Read */, [searchText, searchOptions], true, true));
        };

        Body.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        Body.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ContentControls"])) {
                this.contentControls.handleResult(obj["ContentControls"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Font"])) {
                this.font.handleResult(obj["Font"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["InlinePictures"])) {
                this.inlinePictures.handleResult(obj["InlinePictures"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Paragraphs"])) {
                this.paragraphs.handleResult(obj["Paragraphs"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ParentContentControl"])) {
                this.parentContentControl.handleResult(obj["ParentContentControl"]);
            }
        };

        Body.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Body;
    })(OfficeExtension.ClientObject);
    Word.Body = Body;

    var ContentControl = (function (_super) {
        __extends(ContentControl, _super);
        function ContentControl() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ContentControl.prototype, "contentControls", {
            get: function () {
                if (!this.m_contentControls) {
                    this.m_contentControls = new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ContentControls", true, false));
                }
                return this.m_contentControls;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Word.Font(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "inlinePictures", {
            get: function () {
                if (!this.m_inlinePictures) {
                    this.m_inlinePictures = new Word.InlinePictureCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
                }
                return this.m_inlinePictures;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "paragraphs", {
            get: function () {
                if (!this.m_paragraphs) {
                    this.m_paragraphs = new Word.ParagraphCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
                }
                return this.m_paragraphs;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "parentContentControl", {
            get: function () {
                if (!this.m_parentContentControl) {
                    this.m_parentContentControl = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
                }
                return this.m_parentContentControl;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "appearance", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("appearance", this.m_appearance);
                return this.m_appearance;
            },
            set: function (value) {
                this.m_appearance = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Appearance", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "cannotDelete", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("cannotDelete", this.m_cannotDelete);
                return this.m_cannotDelete;
            },
            set: function (value) {
                this.m_cannotDelete = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "CannotDelete", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "cannotEdit", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("cannotEdit", this.m_cannotEdit);
                return this.m_cannotEdit;
            },
            set: function (value) {
                this.m_cannotEdit = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "CannotEdit", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "color", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("color", this.m_color);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "id", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("id", this.m_id);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "removeWhenEdited", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("removeWhenEdited", this.m_removeWhenEdited);
                return this.m_removeWhenEdited;
            },
            set: function (value) {
                this.m_removeWhenEdited = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "RemoveWhenEdited", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "style", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("style", this.m_style);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "tag", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("tag", this.m_tag);
                return this.m_tag;
            },
            set: function (value) {
                this.m_tag = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Tag", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "title", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("title", this.m_title);
                return this.m_title;
            },
            set: function (value) {
                this.m_title = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Title", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(ContentControl.prototype, "type", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("type", this.m_type);
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        ContentControl.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        ContentControl.prototype.delete = function (keepContent) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Delete", 0 /* Default */, [keepContent]);
        };

        ContentControl.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        ContentControl.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        ContentControl.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        ContentControl.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, true));
        };

        ContentControl.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, true));
        };

        ContentControl.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, true));
        };

        ContentControl.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, true));
        };

        ContentControl.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        ContentControl.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        ContentControl.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Appearance"])) {
                this.m_appearance = obj["Appearance"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["CannotDelete"])) {
                this.m_cannotDelete = obj["CannotDelete"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["CannotEdit"])) {
                this.m_cannotEdit = obj["CannotEdit"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Color"])) {
                this.m_color = obj["Color"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["RemoveWhenEdited"])) {
                this.m_removeWhenEdited = obj["RemoveWhenEdited"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Tag"])) {
                this.m_tag = obj["Tag"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Title"])) {
                this.m_title = obj["Title"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Type"])) {
                this.m_type = obj["Type"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ContentControls"])) {
                this.contentControls.handleResult(obj["ContentControls"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Font"])) {
                this.font.handleResult(obj["Font"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["InlinePictures"])) {
                this.inlinePictures.handleResult(obj["InlinePictures"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Paragraphs"])) {
                this.paragraphs.handleResult(obj["Paragraphs"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ParentContentControl"])) {
                this.parentContentControl.handleResult(obj["ParentContentControl"]);
            }
        };

        ContentControl.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return ContentControl;
    })(OfficeExtension.ClientObject);
    Word.ContentControl = ContentControl;

    var ContentControlCollection = (function (_super) {
        __extends(ContentControlCollection, _super);
        function ContentControlCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ContentControlCollection.prototype, "items", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControlCollection.prototype, "count", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("count", this.m_count);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ContentControlCollection.prototype.getById = function (Id) {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetById", 1 /* Read */, [Id], false, false));
        };

        ContentControlCollection.prototype.getByTag = function (Tag) {
            return new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetByTag", 1 /* Read */, [Tag], true, false));
        };

        ContentControlCollection.prototype.getByTitle = function (Title) {
            return new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetByTitle", 1 /* Read */, [Title], true, false));
        };

        ContentControlCollection.prototype.getItemAt = function (index) {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        ContentControlCollection.prototype.handleResult = function (value) {
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
                    var _item = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };

        ContentControlCollection.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return ContentControlCollection;
    })(OfficeExtension.ClientObject);
    Word.ContentControlCollection = ContentControlCollection;

    var Document = (function (_super) {
        __extends(Document, _super);
        function Document() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Document.prototype, "body", {
            get: function () {
                if (!this.m_body) {
                    this.m_body = new Word.Body(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Body", false, false));
                }
                return this.m_body;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Document.prototype, "contentControls", {
            get: function () {
                if (!this.m_contentControls) {
                    this.m_contentControls = new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ContentControls", true, false));
                }
                return this.m_contentControls;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Document.prototype, "sections", {
            get: function () {
                if (!this.m_sections) {
                    this.m_sections = new Word.SectionCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Sections", true, false));
                }
                return this.m_sections;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Document.prototype, "saved", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("saved", this.m_saved);
                return this.m_saved;
            },
            enumerable: true,
            configurable: true
        });

        Document.prototype.getSelection = function () {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetSelection", 1 /* Read */, [], false, true));
        };

        Document.prototype.save = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Save", 0 /* Default */, []);
        };

        Document.prototype._GetObjectByReferenceId = function (referenceId) {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_GetObjectByReferenceId", 1 /* Read */, [referenceId]);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Document.prototype._GetObjectTypeNameByReferenceId = function (referenceId) {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1 /* Read */, [referenceId]);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Document.prototype._RemoveReference = function (referenceId) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_RemoveReference", 1 /* Read */, [referenceId]);
        };

        Document.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Saved"])) {
                this.m_saved = obj["Saved"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Body"])) {
                this.body.handleResult(obj["Body"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ContentControls"])) {
                this.contentControls.handleResult(obj["ContentControls"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Sections"])) {
                this.sections.handleResult(obj["Sections"]);
            }
        };

        Document.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Document;
    })(OfficeExtension.ClientObject);
    Word.Document = Document;

    var Font = (function (_super) {
        __extends(Font, _super);
        function Font() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Font.prototype, "bold", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("bold", this.m_bold);
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
                OfficeExtension.Utility.throwIfNotLoaded("color", this.m_color);
                return this.m_color;
            },
            set: function (value) {
                this.m_color = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Color", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "doubleStrikeThrough", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("doubleStrikeThrough", this.m_doubleStrikeThrough);
                return this.m_doubleStrikeThrough;
            },
            set: function (value) {
                this.m_doubleStrikeThrough = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "DoubleStrikeThrough", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "highlightColor", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("highlightColor", this.m_highlightColor);
                return this.m_highlightColor;
            },
            set: function (value) {
                this.m_highlightColor = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "HighlightColor", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "italic", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("italic", this.m_italic);
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
                OfficeExtension.Utility.throwIfNotLoaded("name", this.m_name);
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
                OfficeExtension.Utility.throwIfNotLoaded("size", this.m_size);
                return this.m_size;
            },
            set: function (value) {
                this.m_size = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Size", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "strikeThrough", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("strikeThrough", this.m_strikeThrough);
                return this.m_strikeThrough;
            },
            set: function (value) {
                this.m_strikeThrough = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "StrikeThrough", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "subscript", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("subscript", this.m_subscript);
                return this.m_subscript;
            },
            set: function (value) {
                this.m_subscript = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Subscript", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "superscript", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("superscript", this.m_superscript);
                return this.m_superscript;
            },
            set: function (value) {
                this.m_superscript = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Superscript", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "underline", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("underline", this.m_underline);
                return this.m_underline;
            },
            set: function (value) {
                this.m_underline = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Underline", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Font.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Font.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

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

            if (!OfficeExtension.Utility.isUndefined(obj["DoubleStrikeThrough"])) {
                this.m_doubleStrikeThrough = obj["DoubleStrikeThrough"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["HighlightColor"])) {
                this.m_highlightColor = obj["HighlightColor"];
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

            if (!OfficeExtension.Utility.isUndefined(obj["StrikeThrough"])) {
                this.m_strikeThrough = obj["StrikeThrough"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Subscript"])) {
                this.m_subscript = obj["Subscript"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Superscript"])) {
                this.m_superscript = obj["Superscript"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Underline"])) {
                this.m_underline = obj["Underline"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }
        };

        Font.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Font;
    })(OfficeExtension.ClientObject);
    Word.Font = Font;

    var InlinePicture = (function (_super) {
        __extends(InlinePicture, _super);
        function InlinePicture() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(InlinePicture.prototype, "parentContentControl", {
            get: function () {
                if (!this.m_parentContentControl) {
                    this.m_parentContentControl = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
                }
                return this.m_parentContentControl;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(InlinePicture.prototype, "altTextDescription", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("altTextDescription", this.m_altTextDescription);
                return this.m_altTextDescription;
            },
            set: function (value) {
                this.m_altTextDescription = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "AltTextDescription", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "altTextTitle", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("altTextTitle", this.m_altTextTitle);
                return this.m_altTextTitle;
            },
            set: function (value) {
                this.m_altTextTitle = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "AltTextTitle", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "height", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("height", this.m_height);
                return this.m_height;
            },
            set: function (value) {
                this.m_height = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Height", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "hyperlink", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("hyperlink", this.m_hyperlink);
                return this.m_hyperlink;
            },
            set: function (value) {
                this.m_hyperlink = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Hyperlink", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "id", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("id", this.m_id);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(InlinePicture.prototype, "lockAspectRatio", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("lockAspectRatio", this.m_lockAspectRatio);
                return this.m_lockAspectRatio;
            },
            set: function (value) {
                this.m_lockAspectRatio = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LockAspectRatio", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "width", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("width", this.m_width);
                return this.m_width;
            },
            set: function (value) {
                this.m_width = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Width", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(InlinePicture.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        InlinePicture.prototype.getBase64ImageSrc = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetBase64ImageSrc", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        InlinePicture.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, true));
        };

        InlinePicture.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        InlinePicture.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["AltTextDescription"])) {
                this.m_altTextDescription = obj["AltTextDescription"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["AltTextTitle"])) {
                this.m_altTextTitle = obj["AltTextTitle"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Height"])) {
                this.m_height = obj["Height"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Hyperlink"])) {
                this.m_hyperlink = obj["Hyperlink"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LockAspectRatio"])) {
                this.m_lockAspectRatio = obj["LockAspectRatio"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Width"])) {
                this.m_width = obj["Width"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ParentContentControl"])) {
                this.parentContentControl.handleResult(obj["ParentContentControl"]);
            }
        };

        InlinePicture.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return InlinePicture;
    })(OfficeExtension.ClientObject);
    Word.InlinePicture = InlinePicture;

    var InlinePictureCollection = (function (_super) {
        __extends(InlinePictureCollection, _super);
        function InlinePictureCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(InlinePictureCollection.prototype, "items", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(InlinePictureCollection.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        InlinePictureCollection.prototype.getItem = function (index) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        InlinePictureCollection.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        InlinePictureCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };

        InlinePictureCollection.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return InlinePictureCollection;
    })(OfficeExtension.ClientObject);
    Word.InlinePictureCollection = InlinePictureCollection;

    var Paragraph = (function (_super) {
        __extends(Paragraph, _super);
        function Paragraph() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Paragraph.prototype, "contentControls", {
            get: function () {
                if (!this.m_contentControls) {
                    this.m_contentControls = new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ContentControls", true, false));
                }
                return this.m_contentControls;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Paragraph.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Word.Font(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Paragraph.prototype, "inlinePictures", {
            get: function () {
                if (!this.m_inlinePictures) {
                    this.m_inlinePictures = new Word.InlinePictureCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "InlinePictures", true, false));
                }
                return this.m_inlinePictures;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Paragraph.prototype, "parentContentControl", {
            get: function () {
                if (!this.m_parentContentControl) {
                    this.m_parentContentControl = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
                }
                return this.m_parentContentControl;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Paragraph.prototype, "alignment", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("alignment", this.m_alignment);
                return this.m_alignment;
            },
            set: function (value) {
                this.m_alignment = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Alignment", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "firstLineIndent", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("firstLineIndent", this.m_firstLineIndent);
                return this.m_firstLineIndent;
            },
            set: function (value) {
                this.m_firstLineIndent = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "FirstLineIndent", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "leftIndent", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("leftIndent", this.m_leftIndent);
                return this.m_leftIndent;
            },
            set: function (value) {
                this.m_leftIndent = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LeftIndent", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "lineSpacing", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("lineSpacing", this.m_lineSpacing);
                return this.m_lineSpacing;
            },
            set: function (value) {
                this.m_lineSpacing = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LineSpacing", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "lineUnitAfter", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("lineUnitAfter", this.m_lineUnitAfter);
                return this.m_lineUnitAfter;
            },
            set: function (value) {
                this.m_lineUnitAfter = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LineUnitAfter", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "lineUnitBefore", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("lineUnitBefore", this.m_lineUnitBefore);
                return this.m_lineUnitBefore;
            },
            set: function (value) {
                this.m_lineUnitBefore = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "LineUnitBefore", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "outlineLevel", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("outlineLevel", this.m_outlineLevel);
                return this.m_outlineLevel;
            },
            set: function (value) {
                this.m_outlineLevel = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "OutlineLevel", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "rightIndent", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("rightIndent", this.m_rightIndent);
                return this.m_rightIndent;
            },
            set: function (value) {
                this.m_rightIndent = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "RightIndent", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "spaceAfter", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("spaceAfter", this.m_spaceAfter);
                return this.m_spaceAfter;
            },
            set: function (value) {
                this.m_spaceAfter = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "SpaceAfter", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "spaceBefore", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("spaceBefore", this.m_spaceBefore);
                return this.m_spaceBefore;
            },
            set: function (value) {
                this.m_spaceBefore = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "SpaceBefore", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "style", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("style", this.m_style);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Paragraph.prototype, "_Id", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_Id", this.m__Id);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Paragraph.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Paragraph.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        Paragraph.prototype.delete = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };

        Paragraph.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.insertBreak = function (bt, loc) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "InsertBreak", 0 /* Default */, [bt, loc]);
        };

        Paragraph.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, true));
        };

        Paragraph.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, true));
        };

        Paragraph.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, true));
        };

        Paragraph.prototype.insertInlinePictureFromBase64 = function (base64EncodedImage, loc) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertInlinePictureFromBase64", 0 /* Default */, [base64EncodedImage, loc], false, true));
        };

        Paragraph.prototype.insertInlinePictureFromUrl = function (url, loc, linkToFile, saveWithDoc) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertInlinePictureFromUrl", 0 /* Default */, [url, loc, linkToFile, saveWithDoc], false, true));
        };

        Paragraph.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, true));
        };

        Paragraph.prototype.insertParagraph = function (paragraphText, loc) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertParagraph", 0 /* Default */, [paragraphText, loc], false, true));
        };

        Paragraph.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, true));
        };

        Paragraph.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        Paragraph.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        Paragraph.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Alignment"])) {
                this.m_alignment = obj["Alignment"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["FirstLineIndent"])) {
                this.m_firstLineIndent = obj["FirstLineIndent"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LeftIndent"])) {
                this.m_leftIndent = obj["LeftIndent"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LineSpacing"])) {
                this.m_lineSpacing = obj["LineSpacing"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LineUnitAfter"])) {
                this.m_lineUnitAfter = obj["LineUnitAfter"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["LineUnitBefore"])) {
                this.m_lineUnitBefore = obj["LineUnitBefore"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["OutlineLevel"])) {
                this.m_outlineLevel = obj["OutlineLevel"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["RightIndent"])) {
                this.m_rightIndent = obj["RightIndent"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["SpaceAfter"])) {
                this.m_spaceAfter = obj["SpaceAfter"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["SpaceBefore"])) {
                this.m_spaceBefore = obj["SpaceBefore"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ContentControls"])) {
                this.contentControls.handleResult(obj["ContentControls"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Font"])) {
                this.font.handleResult(obj["Font"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["InlinePictures"])) {
                this.inlinePictures.handleResult(obj["InlinePictures"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ParentContentControl"])) {
                this.parentContentControl.handleResult(obj["ParentContentControl"]);
            }
        };

        Paragraph.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Paragraph;
    })(OfficeExtension.ClientObject);
    Word.Paragraph = Paragraph;

    var ParagraphCollection = (function (_super) {
        __extends(ParagraphCollection, _super);
        function ParagraphCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ParagraphCollection.prototype, "items", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ParagraphCollection.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        ParagraphCollection.prototype.getItem = function (index) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        ParagraphCollection.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        ParagraphCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };

        ParagraphCollection.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return ParagraphCollection;
    })(OfficeExtension.ClientObject);
    Word.ParagraphCollection = ParagraphCollection;

    var Range = (function (_super) {
        __extends(Range, _super);
        function Range() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Range.prototype, "contentControls", {
            get: function () {
                if (!this.m_contentControls) {
                    this.m_contentControls = new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ContentControls", true, false));
                }
                return this.m_contentControls;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "font", {
            get: function () {
                if (!this.m_font) {
                    this.m_font = new Word.Font(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Font", false, false));
                }
                return this.m_font;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "paragraphs", {
            get: function () {
                if (!this.m_paragraphs) {
                    this.m_paragraphs = new Word.ParagraphCollection(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Paragraphs", true, false));
                }
                return this.m_paragraphs;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "parentContentControl", {
            get: function () {
                if (!this.m_parentContentControl) {
                    this.m_parentContentControl = new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "ParentContentControl", false, false));
                }
                return this.m_parentContentControl;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "style", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("style", this.m_style);
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(Range.prototype, "_Id", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_Id", this.m__Id);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Range.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Range.prototype.clear = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Clear", 0 /* Default */, []);
        };

        Range.prototype.delete = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };

        Range.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 1 /* Read */, [], false, true));
        };

        Range.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 1 /* Read */, [path, loc], false, true));
        };

        Range.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 1 /* Read */, [html, loc], false, true));
        };

        Range.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 1 /* Read */, [ooxml, loc], false, true));
        };

        Range.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 1 /* Read */, [txt, loc], false, true));
        };

        Range.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        Range.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        Range.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ContentControls"])) {
                this.contentControls.handleResult(obj["ContentControls"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Font"])) {
                this.font.handleResult(obj["Font"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Paragraphs"])) {
                this.paragraphs.handleResult(obj["Paragraphs"]);
            }

            if (!OfficeExtension.Utility.isUndefined(obj["ParentContentControl"])) {
                this.parentContentControl.handleResult(obj["ParentContentControl"]);
            }
        };

        Range.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Range;
    })(OfficeExtension.ClientObject);
    Word.Range = Range;

    var SearchOptions = (function (_super) {
        __extends(SearchOptions, _super);
        function SearchOptions() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(SearchOptions.prototype, "ignorePunct", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("ignorePunct", this.m_ignorePunct);
                return this.m_ignorePunct;
            },
            set: function (value) {
                this.m_ignorePunct = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "IgnorePunct", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "ignoreSpace", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("ignoreSpace", this.m_ignoreSpace);
                return this.m_ignoreSpace;
            },
            set: function (value) {
                this.m_ignoreSpace = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "IgnoreSpace", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchCase", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchCase", this.m_matchCase);
                return this.m_matchCase;
            },
            set: function (value) {
                this.m_matchCase = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchCase", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchPrefix", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchPrefix", this.m_matchPrefix);
                return this.m_matchPrefix;
            },
            set: function (value) {
                this.m_matchPrefix = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchPrefix", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchSoundsLike", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchSoundsLike", this.m_matchSoundsLike);
                return this.m_matchSoundsLike;
            },
            set: function (value) {
                this.m_matchSoundsLike = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchSoundsLike", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchSuffix", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchSuffix", this.m_matchSuffix);
                return this.m_matchSuffix;
            },
            set: function (value) {
                this.m_matchSuffix = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchSuffix", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchWholeWord", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchWholeWord", this.m_matchWholeWord);
                return this.m_matchWholeWord;
            },
            set: function (value) {
                this.m_matchWholeWord = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchWholeWord", value);
            },
            enumerable: true,
            configurable: true
        });


        Object.defineProperty(SearchOptions.prototype, "matchWildCards", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("matchWildCards", this.m_matchWildCards);
                return this.m_matchWildCards;
            },
            set: function (value) {
                this.m_matchWildCards = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "MatchWildCards", value);
            },
            enumerable: true,
            configurable: true
        });


        SearchOptions.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["IgnorePunct"])) {
                this.m_ignorePunct = obj["IgnorePunct"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["IgnoreSpace"])) {
                this.m_ignoreSpace = obj["IgnoreSpace"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchCase"])) {
                this.m_matchCase = obj["MatchCase"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchPrefix"])) {
                this.m_matchPrefix = obj["MatchPrefix"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchSoundsLike"])) {
                this.m_matchSoundsLike = obj["MatchSoundsLike"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchSuffix"])) {
                this.m_matchSuffix = obj["MatchSuffix"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchWholeWord"])) {
                this.m_matchWholeWord = obj["MatchWholeWord"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["MatchWildCards"])) {
                this.m_matchWildCards = obj["MatchWildCards"];
            }
        };

        SearchOptions.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };

        SearchOptions.newObject = function (context) {
            var ret = new Word.SearchOptions(context, OfficeExtension.ObjectPathFactory.createNewObjectObjectPath(context, "Microsoft.WordServices.SearchOptions", false));
            return ret;
        };
        return SearchOptions;
    })(OfficeExtension.ClientObject);
    Word.SearchOptions = SearchOptions;

    var SearchResultCollection = (function (_super) {
        __extends(SearchResultCollection, _super);
        function SearchResultCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(SearchResultCollection.prototype, "items", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(SearchResultCollection.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        SearchResultCollection.prototype.getItem = function (index) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        SearchResultCollection.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        SearchResultCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };

        SearchResultCollection.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return SearchResultCollection;
    })(OfficeExtension.ClientObject);
    Word.SearchResultCollection = SearchResultCollection;

    var Section = (function (_super) {
        __extends(Section, _super);
        function Section() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Section.prototype, "body", {
            get: function () {
                if (!this.m_body) {
                    this.m_body = new Word.Body(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Body", false, false));
                }
                return this.m_body;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Section.prototype, "_Id", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_Id", this.m__Id);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Section.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        Section.prototype.getFooter = function (type) {
            return new Word.Body(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetFooter", 1 /* Read */, [type], false, true));
        };

        Section.prototype.getHeader = function (type) {
            return new Word.Body(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetHeader", 1 /* Read */, [type], false, true));
        };

        Section.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        Section.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Body"])) {
                this.body.handleResult(obj["Body"]);
            }
        };

        Section.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return Section;
    })(OfficeExtension.ClientObject);
    Word.Section = Section;

    var SectionCollection = (function (_super) {
        __extends(SectionCollection, _super);
        function SectionCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(SectionCollection.prototype, "items", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(SectionCollection.prototype, "_ReferenceId", {
            get: function () {
                OfficeExtension.Utility.throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        SectionCollection.prototype.getItem = function (index) {
            return new Word.Section(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        SectionCollection.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };

        SectionCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Word.Section(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };

        SectionCollection.prototype.load = function (option) {
            OfficeExtension.Utility.load(this, option);
            return this;
        };
        return SectionCollection;
    })(OfficeExtension.ClientObject);
    Word.SectionCollection = SectionCollection;

    var ContentControlType = (function () {
        function ContentControlType() {
        }
        ContentControlType.richText = "RichText";
        return ContentControlType;
    })();
    Word.ContentControlType = ContentControlType;

    var ContentControlAppearance = (function () {
        function ContentControlAppearance() {
        }
        ContentControlAppearance.boundingBox = "BoundingBox";

        ContentControlAppearance.tags = "Tags";

        ContentControlAppearance.hidden = "Hidden";
        return ContentControlAppearance;
    })();
    Word.ContentControlAppearance = ContentControlAppearance;

    var UnderlineType = (function () {
        function UnderlineType() {
        }
        UnderlineType.none = "None";

        UnderlineType.single = "Single";

        UnderlineType.word = "Word";

        UnderlineType.double = "Double";

        UnderlineType.dotted = "Dotted";

        UnderlineType.hidden = "Hidden";

        UnderlineType.thick = "Thick";

        UnderlineType.dashLine = "DashLine";

        UnderlineType.dotLine = "DotLine";

        UnderlineType.dotDashLine = "DotDashLine";

        UnderlineType.twoDotDashLine = "TwoDotDashLine";

        UnderlineType.wave = "Wave";
        return UnderlineType;
    })();
    Word.UnderlineType = UnderlineType;

    var BreakType = (function () {
        function BreakType() {
        }
        BreakType.page = "Page";

        BreakType.column = "Column";

        BreakType.next = "Next";

        BreakType.sectionContinuous = "SectionContinuous";

        BreakType.sectionEven = "SectionEven";

        BreakType.sectionOdd = "SectionOdd";

        BreakType.line = "Line";

        BreakType.lineClearLeft = "LineClearLeft";

        BreakType.lineClearRight = "LineClearRight";

        BreakType.textWrapping = "TextWrapping";
        return BreakType;
    })();
    Word.BreakType = BreakType;

    var InsertLocation = (function () {
        function InsertLocation() {
        }
        InsertLocation.before = "Before";

        InsertLocation.after = "After";

        InsertLocation.start = "Start";

        InsertLocation.end = "End";

        InsertLocation.replace = "Replace";
        return InsertLocation;
    })();
    Word.InsertLocation = InsertLocation;

    var Alignment = (function () {
        function Alignment() {
        }
        Alignment.unknown = "Unknown";

        Alignment.left = "Left";

        Alignment.centered = "Centered";

        Alignment.right = "Right";

        Alignment.justified = "Justified";
        return Alignment;
    })();
    Word.Alignment = Alignment;

    var HeaderFooterType = (function () {
        function HeaderFooterType() {
        }
        HeaderFooterType.primary = "Primary";

        HeaderFooterType.firstPage = "FirstPage";

        HeaderFooterType.evenPages = "EvenPages";
        return HeaderFooterType;
    })();
    Word.HeaderFooterType = HeaderFooterType;

    var ErrorCodes = (function () {
        function ErrorCodes() {
        }
        ErrorCodes.generalException = "GeneralException";
        return ErrorCodes;
    })();
    Word.ErrorCodes = ErrorCodes;
})(Word || (Word = {}));
var Word;
(function (Word) {
    var WordClientContext = (function (_super) {
        __extends(WordClientContext, _super);
        function WordClientContext(url) {
            _super.call(this, url);
            this.m_requestExecutor = new OfficeExtension.OfficeJsRequestExecutor();
            this.m_document = new Word.Document(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
            this.rootObject = this.m_document;
        }
        Object.defineProperty(WordClientContext.prototype, "document", {
            get: function () {
                return this.m_document;
            },
            enumerable: true,
            configurable: true
        });
        return WordClientContext;
    })(OfficeExtension.ClientRequestContext);
    Word.WordClientContext = WordClientContext;
})(Word || (Word = {}));

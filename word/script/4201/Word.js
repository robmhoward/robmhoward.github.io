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

        Body.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Body.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Body.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, false));
        };

        Body.prototype.insertParagraph = function (paragraphText, loc) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertParagraph", 0 /* Default */, [paragraphText, loc], false, false));
        };

        Body.prototype.search = function (find) {
            return new Word.SearchResultCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "Search", 1 /* Read */, [find], true, true));
        };

        Body.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
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
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControl.prototype, "removeWhenEdited", {
            get: function () {
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
                return this.m_type;
            },
            enumerable: true,
            configurable: true
        });

        ContentControl.prototype.clearContent = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "ClearContent", 0 /* Default */, []);
        };

        ContentControl.prototype.deleteElement = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteElement", 0 /* Default */, []);
        };

        ContentControl.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        ContentControl.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        ContentControl.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, false));
        };

        ContentControl.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, false));
        };

        ContentControl.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, false));
        };

        ContentControl.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, false));
        };

        ContentControl.prototype.remove = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Remove", 0 /* Default */, []);
        };

        ContentControl.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
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
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ContentControlCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ContentControlCollection.prototype.getById = function (Id) {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetById", 0 /* Default */, [Id], false, false));
        };

        ContentControlCollection.prototype.getByTag = function (Tag) {
            return new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetByTag", 0 /* Default */, [Tag], true, false));
        };

        ContentControlCollection.prototype.getByTitle = function (Title) {
            return new Word.ContentControlCollection(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetByTitle", 0 /* Default */, [Title], true, false));
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

        Object.defineProperty(Document.prototype, "selection", {
            get: function () {
                if (!this.m_selection) {
                    this.m_selection = new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createPropertyObjectPath(this.context, this, "Selection", false, false));
                }
                return this.m_selection;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Document.prototype, "properties", {
            get: function () {
                return this.m_properties;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(Document.prototype, "saved", {
            get: function () {
                return this.m_saved;
            },
            enumerable: true,
            configurable: true
        });

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
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_RemoveReference", 0 /* Default */, [referenceId]);
        };

        Document.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Properties"])) {
                this.m_properties = obj["Properties"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["Saved"])) {
                this.m_saved = obj["Saved"];
            }
        };
        return Document;
    })(OfficeExtension.ClientObject);
    Word.Document = Document;

    var Find = (function (_super) {
        __extends(Find, _super);
        function Find() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Find.prototype, "text", {
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


        Find.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
        };

        Find.newObject = function (context) {
            var ret = new Word.Find(context, OfficeExtension.ObjectPathFactory.createNewObjectObjectPath(context, "Microsoft.WordServices.Find", false));
            return ret;
        };
        return Find;
    })(OfficeExtension.ClientObject);
    Word.Find = Find;

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


        Object.defineProperty(Font.prototype, "doubleStrikeThrough", {
            get: function () {
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


        Object.defineProperty(Font.prototype, "strikeThrough", {
            get: function () {
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
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(InlinePicture.prototype, "lockAspectRatio", {
            get: function () {
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
                return this.m_width;
            },
            set: function (value) {
                this.m_width = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Width", value);
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
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, false));
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
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(InlinePictureCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        InlinePictureCollection.prototype.getItem = function (index) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createIndexerObjectPath(this.context, this, [index]));
        };

        InlinePictureCollection.prototype.getItemAt = function (index) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        InlinePictureCollection.prototype.handleResult = function (value) {
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
                    var _item = new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer(this.context, this, _data[i]));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
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

        Object.defineProperty(Paragraph.prototype, "style", {
            get: function () {
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Paragraph.prototype.clearContent = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "ClearContent", 0 /* Default */, []);
        };

        Paragraph.prototype.deleteElement = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteElement", 0 /* Default */, []);
        };

        Paragraph.prototype.getAlignment = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetAlignment", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getFirstLineIndent = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetFirstLineIndent", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getLeftIndent = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetLeftIndent", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getLineSpacing = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetLineSpacing", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getLineUnitAfter = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetLineUnitAfter", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getLineUnitBefore = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetLineUnitBefore", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getOutlineLevel = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOutlineLevel", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getRightIndent = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetRightIndent", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getSpaceAfter = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetSpaceAfter", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getSpaceBefore = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetSpaceBefore", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Paragraph.prototype.insertBreak = function (bt, loc) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "InsertBreak", 0 /* Default */, [bt, loc]);
        };

        Paragraph.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, false));
        };

        Paragraph.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, false));
        };

        Paragraph.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, false));
        };

        Paragraph.prototype.insertInlinePictureBase64 = function (base64EncodedImage, loc) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertInlinePictureBase64", 0 /* Default */, [base64EncodedImage, loc], false, false));
        };

        Paragraph.prototype.insertInlinePictureUrl = function (url, loc, linkToFile, saveWithDoc) {
            return new Word.InlinePicture(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertInlinePictureUrl", 0 /* Default */, [url, loc, linkToFile, saveWithDoc], false, false));
        };

        Paragraph.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, false));
        };

        Paragraph.prototype.insertParagraph = function (paragraphText, loc) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertParagraph", 0 /* Default */, [paragraphText, loc], false, false));
        };

        Paragraph.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, false));
        };

        Paragraph.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        Paragraph.prototype.setAlignment = function (alignment) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetAlignment", 0 /* Default */, [alignment]);
        };

        Paragraph.prototype.setFirstLineIndent = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetFirstLineIndent", 0 /* Default */, [points]);
        };

        Paragraph.prototype.setLeftIndent = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetLeftIndent", 0 /* Default */, [points]);
        };

        Paragraph.prototype.setLineSpacing = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetLineSpacing", 0 /* Default */, [points]);
        };

        Paragraph.prototype.setLineUnitAfter = function (lines) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetLineUnitAfter", 0 /* Default */, [lines]);
        };

        Paragraph.prototype.setLineUnitBefore = function (lines) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetLineUnitBefore", 0 /* Default */, [lines]);
        };

        Paragraph.prototype.setOutlineLevel = function (level) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetOutlineLevel", 0 /* Default */, [level]);
        };

        Paragraph.prototype.setRightIndent = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetRightIndent", 0 /* Default */, [points]);
        };

        Paragraph.prototype.setSpaceAfter = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetSpaceAfter", 0 /* Default */, [points]);
        };

        Paragraph.prototype.setSpaceBefore = function (points) {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "SetSpaceBefore", 0 /* Default */, [points]);
        };

        Paragraph.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
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
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(ParagraphCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        ParagraphCollection.prototype.getItemAt = function (index) {
            return new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        ParagraphCollection.prototype.handleResult = function (value) {
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
                    var _item = new Word.Paragraph(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
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
                return this.m_style;
            },
            set: function (value) {
                this.m_style = value;
                OfficeExtension.ActionFactory.createSetPropertyAction(this.context, this, "Style", value);
            },
            enumerable: true,
            configurable: true
        });


        Range.prototype.clearContent = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "ClearContent", 0 /* Default */, []);
        };

        Range.prototype.deleteElement = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "DeleteElement", 0 /* Default */, []);
        };

        Range.prototype.getHtml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetHtml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.getOoxml = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetOoxml", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.getText = function () {
            var action = OfficeExtension.ActionFactory.createMethodAction(this.context, this, "GetText", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            this.context.pendingRequest.addActionResultHandler(action, ret);
            return ret;
        };

        Range.prototype.insertContentControl = function () {
            return new Word.ContentControl(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertContentControl", 0 /* Default */, [], false, false));
        };

        Range.prototype.insertFile = function (path, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertFile", 0 /* Default */, [path, loc], false, false));
        };

        Range.prototype.insertHtml = function (html, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertHtml", 0 /* Default */, [html, loc], false, false));
        };

        Range.prototype.insertOoxml = function (ooxml, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertOoxml", 0 /* Default */, [ooxml, loc], false, false));
        };

        Range.prototype.insertText = function (txt, loc) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "InsertText", 0 /* Default */, [txt, loc], false, false));
        };

        Range.prototype.select = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "Select", 0 /* Default */, []);
        };

        Range.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Style"])) {
                this.m_style = obj["Style"];
            }
        };
        return Range;
    })(OfficeExtension.ClientObject);
    Word.Range = Range;

    var RangeCollection = (function (_super) {
        __extends(RangeCollection, _super);
        function RangeCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(RangeCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(RangeCollection.prototype, "_ReferenceId", {
            get: function () {
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        RangeCollection.prototype.getItemAt = function (index) {
            return new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        RangeCollection.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
            if (!OfficeExtension.Utility.isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }

            if (!OfficeExtension.Utility.isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }

            if (!OfficeExtension.Utility.isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return RangeCollection;
    })(OfficeExtension.ClientObject);
    Word.RangeCollection = RangeCollection;

    var SearchResultCollection = (function (_super) {
        __extends(SearchResultCollection, _super);
        function SearchResultCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(SearchResultCollection.prototype, "items", {
            get: function () {
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(SearchResultCollection.prototype, "_ReferenceId", {
            get: function () {
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });

        SearchResultCollection.prototype._KeepReference = function () {
            OfficeExtension.ActionFactory.createMethodAction(this.context, this, "_KeepReference", 0 /* Default */, []);
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
                    var _item = new Word.Range(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
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

        Section.prototype.handleResult = function (value) {
            if (OfficeExtension.Utility.isNullOrUndefined(value))
                return;
            var obj = value;
            OfficeExtension.Utility.fixObjectPathIfNecessary(this, obj);
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
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });

        Object.defineProperty(SectionCollection.prototype, "count", {
            get: function () {
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });

        SectionCollection.prototype.getItemAt = function (index) {
            return new Word.Section(this.context, OfficeExtension.ObjectPathFactory.createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [index], false, false));
        };

        SectionCollection.prototype.handleResult = function (value) {
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
                    var _item = new Word.Section(this.context, OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt(this.context, this, _data[i], i));

                    _item.handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        return SectionCollection;
    })(OfficeExtension.ClientObject);
    Word.SectionCollection = SectionCollection;

    var ContentControlType = (function () {
        function ContentControlType() {
        }
        ContentControlType.unknown = "Unknown";
        ContentControlType.inline = "Inline";
        ContentControlType.paragraph = "Paragraph";
        ContentControlType.cell = "Cell";
        ContentControlType.row = "Row";
        ContentControlType.count = "Count";
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
        Alignment.count = "Count";
        return Alignment;
    })();
    Word.Alignment = Alignment;

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

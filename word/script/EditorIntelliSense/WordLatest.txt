declare module Word {
    class Body extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m__ReferenceId;
        public contentControls : ContentControlCollection;
        public font : Font;
        public inlinePictures : InlinePictureCollection;
        public paragraphs : ParagraphCollection;
        public parentContentControl : ContentControl;
        public style : string;
        public _ReferenceId : string;
        public clear(): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertBreak(bt: string, loc: string): void;
        public insertContentControl(): ContentControl;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertParagraph(paragraphText: string, loc: string): Paragraph;
        public insertText(txt: string, loc: string): Range;
        public search(searchText: string, searchOptions: SearchOptions): SearchResultCollection;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Body;
    }
    class ContentControl extends OfficeExtension.ClientObject {
        private m_appearance;
        private m_cannotDelete;
        private m_cannotEdit;
        private m_color;
        private m_contentControls;
        private m_font;
        private m_id;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        private m_removeWhenEdited;
        private m_style;
        private m_tag;
        private m_title;
        private m_type;
        private m__ReferenceId;
        public contentControls : ContentControlCollection;
        public font : Font;
        public inlinePictures : InlinePictureCollection;
        public paragraphs : ParagraphCollection;
        public parentContentControl : ContentControl;
        public appearance : string;
        public cannotDelete : boolean;
        public cannotEdit : boolean;
        public color : string;
        public id : number;
        public removeWhenEdited : boolean;
        public style : string;
        public tag : string;
        public title : string;
        public type : string;
        public _ReferenceId : string;
        public clear(): void;
        public delete(keepContent: boolean): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertText(txt: string, loc: string): Range;
        public select(): void;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): ContentControl;
    }
    class ContentControlCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : ContentControl[];
        public count : number;
        public getById(Id: number): ContentControl;
        public getByTag(Tag: string): ContentControlCollection;
        public getByTitle(Title: string): ContentControlCollection;
        public getItemAt(index: number): ContentControl;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): ContentControlCollection;
    }
    class Document extends OfficeExtension.ClientObject {
        private m_body;
        private m_contentControls;
        private m_saved;
        private m_sections;
        public body : Body;
        public contentControls : ContentControlCollection;
        public sections : SectionCollection;
        public saved : boolean;
        public getSelection(): Range;
        public save(): void;
        public _GetObjectByReferenceId(referenceId: string): OfficeExtension.ClientResult<any>;
        public _GetObjectTypeNameByReferenceId(referenceId: string): OfficeExtension.ClientResult<string>;
        public _RemoveReference(referenceId: string): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Document;
    }
    class Font extends OfficeExtension.ClientObject {
        private m_bold;
        private m_color;
        private m_doubleStrikeThrough;
        private m_highlightColor;
        private m_italic;
        private m_name;
        private m_size;
        private m_strikeThrough;
        private m_subscript;
        private m_superscript;
        private m_underline;
        private m__ReferenceId;
        public bold : boolean;
        public color : string;
        public doubleStrikeThrough : boolean;
        public highlightColor : string;
        public italic : boolean;
        public name : string;
        public size : number;
        public strikeThrough : boolean;
        public subscript : boolean;
        public superscript : boolean;
        public underline : string;
        public _ReferenceId : string;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Font;
    }
    class InlinePicture extends OfficeExtension.ClientObject {
        private m_altTextDescription;
        private m_altTextTitle;
        private m_height;
        private m_hyperlink;
        private m_id;
        private m_lockAspectRatio;
        private m_parentContentControl;
        private m_width;
        private m__ReferenceId;
        public parentContentControl : ContentControl;
        public altTextDescription : string;
        public altTextTitle : string;
        public height : number;
        public hyperlink : string;
        public id : number;
        public lockAspectRatio : boolean;
        public width : number;
        public _ReferenceId : string;
        public getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        public insertContentControl(): ContentControl;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): InlinePicture;
    }
    class InlinePictureCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        public items : InlinePicture[];
        public _ReferenceId : string;
        public getItem(index: any): InlinePicture;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): InlinePictureCollection;
    }
    class Paragraph extends OfficeExtension.ClientObject {
        private m_alignment;
        private m_contentControls;
        private m_firstLineIndent;
        private m_font;
        private m_inlinePictures;
        private m_leftIndent;
        private m_lineSpacing;
        private m_lineUnitAfter;
        private m_lineUnitBefore;
        private m_outlineLevel;
        private m_parentContentControl;
        private m_rightIndent;
        private m_spaceAfter;
        private m_spaceBefore;
        private m_style;
        private m__Id;
        private m__ReferenceId;
        public contentControls : ContentControlCollection;
        public font : Font;
        public inlinePictures : InlinePictureCollection;
        public parentContentControl : ContentControl;
        public alignment : string;
        public firstLineIndent : number;
        public leftIndent : number;
        public lineSpacing : number;
        public lineUnitAfter : number;
        public lineUnitBefore : number;
        public outlineLevel : number;
        public rightIndent : number;
        public spaceAfter : number;
        public spaceBefore : number;
        public style : string;
        public _Id : number;
        public _ReferenceId : string;
        public clear(): void;
        public delete(): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertBreak(bt: string, loc: string): void;
        public insertContentControl(): ContentControl;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertInlinePictureFromBase64(base64EncodedImage: string, loc: string): InlinePicture;
        public insertInlinePictureFromUrl(url: string, loc: string, linkToFile: boolean, saveWithDoc: boolean): InlinePicture;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertParagraph(paragraphText: string, loc: string): Paragraph;
        public insertText(txt: string, loc: string): Range;
        public select(): void;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Paragraph;
    }
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        public items : Paragraph[];
        public _ReferenceId : string;
        public getItem(index: any): Paragraph;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): ParagraphCollection;
    }
    class Range extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_paragraphs;
        private m_parentContentControl;
        private m_style;
        private m__Id;
        private m__ReferenceId;
        public contentControls : ContentControlCollection;
        public font : Font;
        public paragraphs : ParagraphCollection;
        public parentContentControl : ContentControl;
        public style : string;
        public _Id : number;
        public _ReferenceId : string;
        public clear(): void;
        public delete(): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertContentControl(): ContentControl;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertText(txt: string, loc: string): Range;
        public select(): void;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Range;
    }
    class SearchOptions extends OfficeExtension.ClientObject {
        private m_ignorePunct;
        private m_ignoreSpace;
        private m_matchCase;
        private m_matchPrefix;
        private m_matchSoundsLike;
        private m_matchSuffix;
        private m_matchWholeWord;
        private m_matchWildCards;
        public ignorePunct : boolean;
        public ignoreSpace : boolean;
        public matchCase : boolean;
        public matchPrefix : boolean;
        public matchSoundsLike : boolean;
        public matchSuffix : boolean;
        public matchWholeWord : boolean;
        public matchWildCards : boolean;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): SearchOptions;
        static newObject(context: OfficeExtension.ClientRequestContext): SearchOptions;
    }
    class SearchResultCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        public items : Range[];
        public _ReferenceId : string;
        public getItem(index: any): Range;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): SearchResultCollection;
    }
    class Section extends OfficeExtension.ClientObject {
        private m_body;
        private m__Id;
        private m__ReferenceId;
        public body : Body;
        public _Id : number;
        public _ReferenceId : string;
        public getFooter(type: string): Body;
        public getHeader(type: string): Body;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): Section;
    }
    class SectionCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        public items : Section[];
        public _ReferenceId : string;
        public getItem(index: any): Section;
        public _KeepReference(): void;
        public handleResult(value: any): void;
        public load(option?: OfficeExtension.LoadOption): SectionCollection;
    }
    class ContentControlType {
        static richText: string;
    }
    class ContentControlAppearance {
        static boundingBox: string;
        static tags: string;
        static hidden: string;
    }
    class UnderlineType {
        static none: string;
        static single: string;
        static word: string;
        static double: string;
        static dotted: string;
        static hidden: string;
        static thick: string;
        static dashLine: string;
        static dotLine: string;
        static dotDashLine: string;
        static twoDotDashLine: string;
        static wave: string;
    }
    class BreakType {
        static page: string;
        static column: string;
        static next: string;
        static sectionContinuous: string;
        static sectionEven: string;
        static sectionOdd: string;
        static line: string;
        static lineClearLeft: string;
        static lineClearRight: string;
        static textWrapping: string;
    }
    class InsertLocation {
        static before: string;
        static after: string;
        static start: string;
        static end: string;
        static replace: string;
    }
    class Alignment {
        static unknown: string;
        static left: string;
        static centered: string;
        static right: string;
        static justified: string;
    }
    class HeaderFooterType {
        static primary: string;
        static firstPage: string;
        static evenPages: string;
    }
    class ErrorCodes {
        static generalException: string;
    }
}
declare module Word {
    class WordClientContext extends OfficeExtension.ClientRequestContext {
        private m_document;
        constructor(url?: string);
        public document : Document;
    }
}

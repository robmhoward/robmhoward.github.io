declare module Word {
    class Body extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_inlinePictures;
        private m_paragraphs;
        private m_parentContentControl;
        public contentControls : ContentControlCollection;
        public inlinePictures : InlinePictureCollection;
        public paragraphs : ParagraphCollection;
        public parentContentControl : ContentControl;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public insertContentControl(): ContentControl;
        public insertParagraph(paragraphText: string, loc: string): Paragraph;
        public search(find: Find): SearchResultCollection;
        public handleResult(value: any): void;
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
        public clearContent(): void;
        public deleteElement(): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertText(txt: string, loc: string): Range;
        public remove(): void;
        public select(): void;
        public handleResult(value: any): void;
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
    }
    class Document extends OfficeExtension.ClientObject {
        private m_body;
        private m_contentControls;
        private m_properties;
        private m_saved;
        private m_sections;
        private m_selection;
        public body : Body;
        public contentControls : ContentControlCollection;
        public sections : SectionCollection;
        public selection : Range;
        public properties : string;
        public saved : boolean;
        public save(): void;
        public _GetObjectByReferenceId(referenceId: string): OfficeExtension.ClientResult<any>;
        public _GetObjectTypeNameByReferenceId(referenceId: string): OfficeExtension.ClientResult<string>;
        public _RemoveReference(referenceId: string): void;
        public handleResult(value: any): void;
    }
    class Find extends OfficeExtension.ClientObject {
        private m_text;
        public text : string;
        public handleResult(value: any): void;
        static newObject(context: OfficeExtension.ClientRequestContext): Find;
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
        public handleResult(value: any): void;
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
        public parentContentControl : ContentControl;
        public altTextDescription : string;
        public altTextTitle : string;
        public height : number;
        public hyperlink : string;
        public id : number;
        public lockAspectRatio : boolean;
        public width : number;
        public getBase64ImageSrc(): OfficeExtension.ClientResult<string>;
        public insertContentControl(): ContentControl;
        public handleResult(value: any): void;
    }
    class InlinePictureCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : InlinePicture[];
        public count : number;
        public getItem(index: any): InlinePicture;
        public getItemAt(index: number): InlinePicture;
        public handleResult(value: any): void;
    }
    class Paragraph extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_parentContentControl;
        private m_style;
        public contentControls : ContentControlCollection;
        public font : Font;
        public parentContentControl : ContentControl;
        public style : string;
        public clearContent(): void;
        public deleteElement(): void;
        public getAlignment(): OfficeExtension.ClientResult<string>;
        public getFirstLineIndent(): OfficeExtension.ClientResult<number>;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getLeftIndent(): OfficeExtension.ClientResult<number>;
        public getLineSpacing(): OfficeExtension.ClientResult<number>;
        public getLineUnitAfter(): OfficeExtension.ClientResult<number>;
        public getLineUnitBefore(): OfficeExtension.ClientResult<number>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getOutlineLevel(): OfficeExtension.ClientResult<number>;
        public getRightIndent(): OfficeExtension.ClientResult<number>;
        public getSpaceAfter(): OfficeExtension.ClientResult<number>;
        public getSpaceBefore(): OfficeExtension.ClientResult<number>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertBreak(bt: string, loc: string): void;
        public insertContentControl(): ContentControl;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertInlinePictureBase64(base64EncodedImage: string, loc: string): InlinePicture;
        public insertInlinePictureUrl(url: string, loc: string, linkToFile: boolean, saveWithDoc: boolean): InlinePicture;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertParagraph(paragraphText: string, loc: string): Paragraph;
        public insertText(txt: string, loc: string): Range;
        public select(): void;
        public setAlignment(alignment: string): void;
        public setFirstLineIndent(points: number): void;
        public setLeftIndent(points: number): void;
        public setLineSpacing(points: number): void;
        public setLineUnitAfter(lines: number): void;
        public setLineUnitBefore(lines: number): void;
        public setOutlineLevel(level: number): void;
        public setRightIndent(points: number): void;
        public setSpaceAfter(points: number): void;
        public setSpaceBefore(points: number): void;
        public handleResult(value: any): void;
    }
    class ParagraphCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : Paragraph[];
        public count : number;
        public getItemAt(index: number): Paragraph;
        public handleResult(value: any): void;
    }
    class Range extends OfficeExtension.ClientObject {
        private m_contentControls;
        private m_font;
        private m_parentContentControl;
        private m_style;
        public contentControls : ContentControlCollection;
        public font : Font;
        public parentContentControl : ContentControl;
        public style : string;
        public clearContent(): void;
        public deleteElement(): void;
        public getHtml(): OfficeExtension.ClientResult<string>;
        public getOoxml(): OfficeExtension.ClientResult<string>;
        public getText(): OfficeExtension.ClientResult<string>;
        public insertContentControl(): ContentControl;
        public insertFile(path: string, loc: string): Range;
        public insertHtml(html: string, loc: string): Range;
        public insertOoxml(ooxml: string, loc: string): Range;
        public insertText(txt: string, loc: string): Range;
        public select(): void;
        public handleResult(value: any): void;
    }
    class RangeCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__ReferenceId;
        private m__items;
        public items : Range[];
        public count : number;
        public _ReferenceId : string;
        public getItemAt(index: number): Range;
        public handleResult(value: any): void;
    }
    class SearchResultCollection extends OfficeExtension.ClientObject {
        private m__ReferenceId;
        private m__items;
        public items : Range[];
        public _ReferenceId : string;
        public _KeepReference(): void;
        public handleResult(value: any): void;
    }
    class Section extends OfficeExtension.ClientObject {
        private m_body;
        public body : Body;
        public handleResult(value: any): void;
    }
    class SectionCollection extends OfficeExtension.ClientObject {
        private m_count;
        private m__items;
        public items : Section[];
        public count : number;
        public getItemAt(index: number): Section;
        public handleResult(value: any): void;
    }
    class ContentControlType {
        static unknown: string;
        static inline: string;
        static paragraph: string;
        static cell: string;
        static row: string;
        static count: string;
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
        static count: string;
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

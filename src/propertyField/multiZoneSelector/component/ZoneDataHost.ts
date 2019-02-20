import { IZoneData, 
    ZoneDataType, 
    IContentData, 
    IVideoData, 
    getContentDataDefaultValue, 
    getArticleDataDefaultValue, 
    getVideoDataDefaultValue } from "../IPropertyPaneMultiZoneSelector";
import { IPropertyFieldNewsSelectorData } from "../../newsSelector";

export class ZoneDataHost {
    private dataType: ZoneDataType;
    private contentData: IContentData;
    private videoData: IVideoData;
    private articleData: IPropertyFieldNewsSelectorData;

    constructor(data?: IZoneData) {
        this.dataType = ZoneDataType.Content;
        this.contentData = getContentDataDefaultValue();
        this.articleData = getArticleDataDefaultValue();
        this.videoData = getVideoDataDefaultValue();
        this.setData(data);
    }

    public setData(data: IZoneData): void {
        if (typeof data !== "undefined") {
            if (typeof data.type !== "undefined") {
                this.dataType = data.type;
            }
            switch (this.dataType) {
                case ZoneDataType.Article: {
                    if (typeof data.data !== "undefined") {
                        this.articleData = data.data as IPropertyFieldNewsSelectorData;
                    }
                    else {
                        this.articleData = getArticleDataDefaultValue();
                    }
                    break;
                }
                case ZoneDataType.Video: {
                    if (typeof data.data !== "undefined") {
                        this.videoData = data.data as IVideoData;
                    }
                    else {
                        this.videoData = getVideoDataDefaultValue();
                    }
                    break;
                }
                default: {
                    if (typeof data.data !== "undefined") {
                        this.contentData = data.data as IContentData;
                    }
                    else {
                        this.contentData = getContentDataDefaultValue();
                    }
                    break;
                }
            }
        }
    }

    public getData(): IContentData | IVideoData | IPropertyFieldNewsSelectorData {
        return (this.dataType === ZoneDataType.Video) ? this.videoData :
            ((this.dataType === ZoneDataType.Article) ? this.articleData : this.contentData);
    }

    public getType(): ZoneDataType {
        return this.dataType;
    }

    public setZoneType(typeValue: string | number): void {
        const value: number = parseInt(typeValue.toString());
        if (value === ZoneDataType.Video) {
            this.dataType = ZoneDataType.Video;
            if (typeof this.videoData === "undefined" || this.videoData === null) {
                this.videoData = getVideoDataDefaultValue();
            }
        } else {
            if (value === ZoneDataType.Article) {
                this.dataType = ZoneDataType.Article;
                if (typeof this.articleData === "undefined" || this.articleData === null) {
                    this.articleData = getArticleDataDefaultValue();
                }
            } else {
                this.dataType = ZoneDataType.Content;
                if (typeof this.contentData === "undefined" || this.contentData === null) {
                    this.contentData = getContentDataDefaultValue();
                }
            }
        }
    }
}

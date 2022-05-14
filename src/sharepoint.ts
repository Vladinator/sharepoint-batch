import { RequestOptions, SharePointOptions, SharePointParams, SPWeb } from './types';
import { isObject, isString, extend, toParams } from './utils';
import { RequestJson } from './request';

export default class SharePointBatch {

    static GetSharePointOptions(): SharePointOptions | undefined {

        //@ts-expect-error
        const context: any = window._spPageContextInfo;

        //@ts-expect-error
        const digest: Function = window.GetRequestDigest;

        if (!isObject(context) || typeof digest !== 'function')
            return;

        return {
            url: context.webAbsoluteUrl,
            digest: digest(),
        };

    }

    static CreateGUID(): string {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c: string) => {
            const r = Math.random() * 16 | 0;
            return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
        });
    }

    _options: SharePointOptions;

    constructor(options: SharePointOptions) {

        this._options = options;

    }

    GetRequestOptions(url: string | RequestOptions): RequestOptions {

        const options: RequestOptions = {
            method: 'GET',
            url: '',
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': 'application/json;odata=verbose',
            },
        };

        if (isObject(url)) {

            //@ts-expect-error
            if (isObject(url.headers)) {
                //@ts-expect-error
                extend(options.headers, url.headers);
            }

            extend(options, url);

        } else if (isString(url)) {

            //@ts-expect-error
            options.url = url;

        }

        return options;

    }

    async QueryEndpoint<T>(url: string): Promise<T | undefined> {
        const options = this.GetRequestOptions(url);
        const payload = await RequestJson(options);
        if (payload)
            return payload.d;
    }

    async GetWeb(params?: SharePointParams): Promise<SPWeb | undefined> {
        return await this.QueryEndpoint<SPWeb>(`/_api/Web${toParams(params)}`);
    }

}

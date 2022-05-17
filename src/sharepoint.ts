import { RequestMethods, RequestOptions, SharePointOptions, BatchJobOptions, BatchJobHeader, SharePointBatchJobResponse, ResponseParserPayload, SharePointBatchResponse, SharePointBatchResponseSuccess, SharePointBatchResponseError } from './types';
import { isArray, isObject, isString, extend, toParams, createGUID, safeCall } from './utils';
import { Request } from './request';

const FallbackBatchJobOptions: BatchJobOptions = {
    method: 'GET',
    url: '',
};

const FallbackRequestOptions: RequestOptions = {
    method: 'POST',
    url: '',
    headers: {},
};

/** @internal */
class ResponseParser {

    static Level = {
        UNKNOWN: 0,
        HEADERS: 1,
        REQUEST: 2,
        REQUEST_HEADERS: 3,
        REQUEST_BODY: 4,
        EOF: 5,
    };

    static LineSeparator = /\r\n/;
    static HeaderKVSeparator = /:/;
    static HeaderKVSeparatorChar = ':';
    static BatchResponse1 = /^--batchresponse_.+--$/i;
    static BatchResponse2 = /^--batchresponse_.+$/i;
    static BatchResponse3 = /^HTTP\/1\.1\s+(\d+)\s+(.+)$/i;

    static Parse(raw: string): ResponseParserPayload[] | any {

        if (!isString(raw))
            return;

        try {
            return JSON.parse(raw);
        } catch (ex: any) {
        }

        const parseLevels = ResponseParser.Level;
        const lines = raw.split(ResponseParser.LineSeparator);
        const results: ResponseParserPayload[] = [];

        let temp: ResponseParserPayload | null = null;
        let cwo: ResponseParserPayload | null = null;
        let level = parseLevels.UNKNOWN;

        for (const line of lines) {

            if (ResponseParser.BatchResponse1.test(line)) {

                if (temp) {
                    temp.data = this.Parse(temp.data);
                    results.push(temp);
                    temp = null;
                }

                level = parseLevels.EOF;
                break;

            } else if (ResponseParser.BatchResponse2.test(line)) {

                if (temp) {
                    temp.data = this.Parse(temp.data);
                    results.push(temp);
                }

                temp = {
                    headers: {},
                    http: { status: 0, statusText: '' },
                    ok: false,
                    data: null,
                };

                cwo = temp;
                level = parseLevels.HEADERS;

            } else if (level === parseLevels.REQUEST_BODY) {

                if (cwo) {

                    if (cwo.data == null)
                        cwo.data = '';

                    cwo.data += line;

                }

            } else if (ResponseParser.BatchResponse3.test(line)) {

                if (level === parseLevels.REQUEST) {

                    if (cwo) {

                        const http = line.match(ResponseParser.BatchResponse3);
                        cwo.http.status = http && parseInt(http[1], 10) || 0;
                        cwo.http.statusText = http && http[2] || '';
                        cwo.ok = (cwo.http.status / 100 | 0) === 2;
                        level = parseLevels.REQUEST_HEADERS;

                    }

                }

            } else if (/^.+:\s*.+$/i.test(line)) {

                if (level === parseLevels.HEADERS || level === parseLevels.REQUEST_HEADERS) {

                    const parts = line.split(ResponseParser.HeaderKVSeparator);

                    if (parts) {

                        const rawHeader = parts.shift();
                        const key = rawHeader && rawHeader.trim();

                        if (key && cwo)
                            cwo.headers[key] = parts.join(ResponseParser.HeaderKVSeparatorChar).trim();

                    }

                }

            } else if (/^[\s\r\n]*$/i.test(line)) {

                switch (level) {

                    case parseLevels.HEADERS:
                        level = parseLevels.REQUEST;
                        break;

                    case parseLevels.REQUEST:
                        level = parseLevels.REQUEST_HEADERS;
                        break;

                    case parseLevels.REQUEST_HEADERS:
                        level = parseLevels.REQUEST_BODY;
                        break;

                }

            }

        }

        if (!results.length) {
            const safeParse = ResponseParser.SafeParse(raw);
            if (safeParse && safeParse !== raw)
                return safeParse;
        }

        return results;

    }

    static SafeParse(raw: any): ResponseParserPayload | any {

        if (!isString(raw))
            return raw;

        try {
            return JSON.parse(raw);
        } catch (ex: any) {
        }

        let dom: DOMParser | undefined;

        try {
            dom = new DOMParser();
        } catch (ex: any) {
        }

        if (!dom)
            return raw;

        let doc: Document | undefined;

        try {
            doc = dom.parseFromString(raw, 'text/xml');
        } catch (ex: any) {
        }

        if (!doc)
            return raw;

        const error = doc.querySelector('error');

        if (!error)
            return doc;

        const code = error.querySelector('code');
        const message = error.querySelector('message');

        const codeText = code ? code.innerHTML : 0;
        const messageText = message ? message.innerHTML : '';
        const errorText = `${codeText}: ${messageText}`;

        return errorText;

    }

}

/**
 * Individual queries are bundled into "changeset" entries.
 */
export class Changeset {

    /** @internal */
    #options: BatchJobOptions;

    /** @internal */
    #responsePayload: ResponseParserPayload | string | undefined;

    constructor(options: BatchJobOptions) {
        //@ts-expect-error
        this.#options = extend({}, FallbackBatchJobOptions, options);
    }

    getOptions() {
        return this.#options;
    }

    /**
     * Returns the `GET`, `POST`, `...` method.
     * @returns Request method.
     */
    getMethod(): RequestMethods {
        return this.#options.method;
    }

    /**
     * Returns the full request URL.
     * @returns Request URL with any optional params.
     */
    getUrl() {
        return `${this.#options.url}${toParams(this.#options.params)}`;
    }

    /**
     * Returns the headers object for the request.
     * @returns Request header array.
     */
    getHeaders(): BatchJobHeader[] {

        const headers = this.#options.headers;
        const results: BatchJobHeader[] = [];

        if (isObject(headers)) {
            for (const header in headers) {
                results.push({ key: header, value: headers[header] });
            }
        }

        return results;

    }

    /** @internal */
    getPayload(): BodyInit | null | undefined {
        return this.#options.body;
    }

    /**
     * Returns the payload from the server once the changeset is queried and processed.
     * @returns The raw data the server responded with from the request.
     */
    getResponsePayload() {
        return this.#responsePayload;
    }

    /** @internal */
    setResponsePayload(payload: ResponseParserPayload | string | undefined) {
        this.#responsePayload = payload;
    }

}

/**
 * Bundles of Changeset objects are lumped into "batch job" entries.
 * @internal
 */
class BatchJob {

    static NumMaxChangesets = 100;

    /** @internal */
    #options: BatchJobOptions;

    /** @internal */
    #changesets: Changeset[];

    constructor(options: BatchJobOptions) {
        //@ts-expect-error
        this.#options = extend({}, FallbackBatchJobOptions, options);
        this.#changesets = [];
    }

    getOptions() {
        return this.#options;
    }

    getChangesets() {
        return this.#changesets;
    }

    /** @ignore */
    isChangesetsFull() {
        return this.#changesets.length >= BatchJob.NumMaxChangesets;
    }

    /**
     * Append a changeset to the changeset queue.
     * @param changeset An object of `Changeset`.
     * @returns Successfull additions return the `index` in the queue otherwise `-1`.
     */
    addChangeset(changeset: Changeset): number {
        if (changeset instanceof Changeset) {
            const options = changeset.getOptions();
            if (options.url[0] === '/')
                options.url = `${this.#options.url}${options.url}`;
            return this.#changesets.push(changeset) - 1;
        }
        return -1;
    }

    /** @internal */
    getPayload(guid: string): string {

        const data = [];

        for (const changeset of this.#changesets) {

            const method = changeset.getMethod();
            const boundary = method === 'GET' ? null : `changeset_${createGUID()}`;

            data.push(`--batch_${guid}`);

            if (boundary) {
                data.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
            } else {
                data.push('Content-Type: application/http');
            }

            data.push('Content-Transfer-Encoding: binary');
            data.push('');

            if (method === 'GET') {

                data.push(`${method} ${changeset.getUrl()} HTTP/1.1`);
                data.push('Accept: application/json;odata=verbose');

                for (const header of changeset.getHeaders()) {
                    data.push(`${header.key}: ${header.value}`);
                }

                data.push('');

            } else {

                data.push(`--${boundary}`);
                data.push('Content-Type: application/http');
                data.push('Content-Transfer-Encoding: binary');
                data.push('');

                data.push(`${method} ${changeset.getUrl()} HTTP/1.1`);
                data.push('Accept: application/json;odata=verbose');
                data.push('Content-Type: application/json;odata=verbose');

                for (const header of changeset.getHeaders()) {
                    data.push(`${header.key}: ${header.value}`);
                }

                data.push('');

                const changesetPayload = changeset.getPayload();

                if (changesetPayload) {

                    data.push(JSON.stringify(changesetPayload));
                    data.push('');

                }

            }

            if (boundary) {
                data.push(`--${boundary}--`);
            }

        }

        return data.join('\r\n');

    }

    /** @internal */
    #getSendOptions(batch: SharePointBatch, options?: RequestOptions): RequestOptions {

        const batchOptions = batch.getOptions();

        //@ts-expect-error
        const fallback: RequestOptions = extend({}, FallbackRequestOptions, batchOptions, options);

        const guid = createGUID();

        fallback.url = `${batchOptions.url}/_api/$batch`;

        if (fallback.headers) {
            extend(fallback.headers, {
                'Content-Type': `multipart/mixed; boundary="batch_${guid}"`,
                'X-RequestDigest': batchOptions.digest,
            });
        }

        fallback.body = `${this.getPayload(guid)}\r\n--batch_${guid}--`;

        return fallback;

    }

    /** @internal */
    #processResponsePayload(changeset: Changeset, payload: ResponseParserPayload | string | undefined) {

        changeset.setResponsePayload(payload);

        const isPayload = isObject(payload);
        const safePayload = isPayload ? payload as ResponseParserPayload : null;
        const ok = (safePayload ? safePayload.http.status / 100 | 0 : 0) === 2;

        if (safePayload)
            safePayload.changeset = changeset;

        const options = changeset.getOptions();

        if (ok)
            safeCall(options, 'fail', changeset, payload);
        else
            safeCall(options, 'done', changeset, payload);

    }

    /** @internal */
    async send(batch: SharePointBatch, options?: RequestOptions): Promise<SharePointBatchJobResponse> {

        const changesets = this.#changesets;
        changesets.forEach(changeset => safeCall(changeset.getOptions(), 'before', changeset));

        const safeError = (...args: any): undefined => {
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'fail', changeset, ...args));
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'finally', changeset));
            return;
        };

        const safeFinally = () => {
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'finally', changeset));
        };

        const fallback: RequestOptions = this.#getSendOptions(batch, options);

        //@ts-expect-error
        const backup: RequestOptions = extend({}, fallback);

        let delayedDone = false;
        let delayedFail = false;
        fallback.done = () => delayedDone = true;
        fallback.fail = () => delayedFail = true;

        const response = await Request(fallback);

        if (!response) {
            if (delayedFail)
                safeCall(backup, 'fail', response);
            safeError();
            return;
        }

        const payload = await response.text();

        if (!payload || !response.ok) {
            const safePayload = ResponseParser.SafeParse(payload);
            if (delayedFail)
                safeCall(backup, 'fail', response, safePayload, response.status, response.statusText);
            safeError(safePayload, response.status, response.statusText);
            return;
        }

        const parsed = ResponseParser.Parse(payload);

        if (!isArray(parsed)) {
            changesets.forEach(changeset => this.#processResponsePayload(changeset, payload));
            if (delayedDone)
                safeCall(backup, 'done', response, payload);
            safeFinally();
            return payload;
        }

        const changesetPayloads = parsed as ResponseParserPayload[];
        changesets.forEach((changeset, index) => this.#processResponsePayload(changeset, changesetPayloads[index]));
        if (delayedDone)
            safeCall(backup, 'done', response, changesetPayloads);
        safeFinally();
        return changesetPayloads;

    }

}

/**
 * The library entry point class.
 */
export class SharePointBatch {

    /**
     * Utility function to extract data from the `window` properties `_spPageContextInfo` and `GetRequestDigest`.
     * @returns If possible it returns a `SharePointOptions` object otherwise nothing.
     */
    static GetSharePointOptions(): SharePointOptions | undefined {

        const win: any = window;
        const context: any = win._spPageContextInfo;
        const getDigest: any = win.GetRequestDigest;

        if (!isObject(context) || typeof getDigest !== 'function')
            return;

        let url: any = context.webAbsoluteUrl;
        let digest: any;

        try {
            digest = getDigest();
        } catch (ex: any) {
        }

        if (!isString(url))
            url = '';

        if (!isString(digest))
            digest = '';

        return { url, digest };

    }

    /** @internal */
    #options: SharePointOptions;

    /** @internal */
    #jobs: BatchJob[];

    /** @internal */
    #job: BatchJob | undefined;

    constructor(options: SharePointOptions) {
        this.#options = options;
        this.#jobs = [];
    }

    getOptions() {
        return this.#options;
    }

    /** @internal */
    #appendNewJob(options?: BatchJobOptions): BatchJob {

        //@ts-expect-error
        const fallback: BatchJobOptions = extend({}, FallbackBatchJobOptions, this.#options, options);

        this.#job = new BatchJob(fallback);
        this.#jobs.push(this.#job);

        return this.#job;

    }

    /** @internal */
    #getActiveJob(): BatchJob {
        if (!this.#job)
            return this.#appendNewJob();
        return this.#job;
    }

    /**
     * Append a changeset to the batch queue.
     * @param changeset Object instance of `Changeset`.
     * @returns `true` if the changeset was added otherwise `false`.
     */
    addChangeset(changeset: Changeset): boolean {
        let job = this.#getActiveJob();
        if (job.isChangesetsFull())
            job = this.#appendNewJob();
        return job.addChangeset(changeset) > -1;
    }

    /**
     * Process the batch queue. Supports `await` but the results is all the returned data for each changeset.
     * 
     * You can assign the `done` and `fail` to the optional argument `options` and they will be called based on estimation given the results of the request.
     * Since there could be multiple jobs involved, each job will potentially succeed or fail, depending on the changesets submitted to the server.
     * You'll need to inspect each changeset result and see if it succeeded or failed, if the `SharePointBatchResponse` object is present check the `ok` property for the http status.
     * @param options Optional `RequestOptions` object.
     * @returns `Promise` that returns an array of `SharePointBatchResponse` but in case of errors the array item will be `undefined` or a `string` whenever available.
     */
    async send(options?: RequestOptions): Promise<SharePointBatchResponse> {

        //@ts-expect-error
        const fallback: RequestOptions = extend({}, options);

        //@ts-expect-error
        const backup: RequestOptions = extend({}, fallback);

        let delayedDone = false;
        let delayedFail = false;
        let delayedArgs: any[] = [];
        fallback.done = (...args: any[]) => (delayedDone = true, delayedArgs = args);
        fallback.fail = (...args: any[]) => (delayedFail = true, delayedArgs = args);
        fallback.finally = () => {};

        const results: SharePointBatchJobResponse[] = [];
        let success = 0;
        let fail = 0;

        for (const job of this.#jobs) {

            const result = await job.send(this, fallback);

            if (isArray(result)) {

                const safeResult: any = result;
                success += safeResult.length;
                results.push(...safeResult);

            } else {

                let count = job.getChangesets().length;
                fail += count;
                while (count-- > 0)
                    results.push(result);

            }

        }

        // if we have all fails we override the done into a fail
        if (delayedDone && !delayedFail && !success && fail) {
            delayedDone = false;
            delayedFail = true;
        }

        // if we have no done or fail it means the queue was empty
        if (!delayedDone && !delayedFail) {
            delayedDone = true;
            delayedArgs[0] = backup;
        }

        if (delayedDone) {

            safeCall(backup, 'done', delayedArgs[1], results);
            safeCall(backup, 'finally', delayedArgs[1], results);
            return { success: true, ok: true, results } as SharePointBatchResponseSuccess;

        } else {

            safeCall(backup, 'fail', delayedArgs[1], delayedArgs[2], delayedArgs[3], delayedArgs[4]);
            safeCall(backup, 'finally', delayedArgs[1], delayedArgs[2], delayedArgs[3], delayedArgs[4]);
            return { error: true, ok: false, results: delayedArgs[2] } as SharePointBatchResponseError;

        }

    }

}

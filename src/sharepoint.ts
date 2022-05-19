import { RequestMethods, RequestOptions, SharePointOptions, BatchJobOptions, BatchJobHeader, SharePointBatchJobResponse, ResponseParserPayload, SharePointBatchResponse, SharePointBatchResponseSuccess, SharePointBatchResponseError, SharePointBatchOptions } from './types';
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

    private static Level = {
        UNKNOWN: 0,
        HEADERS: 1,
        REQUEST: 2,
        REQUEST_HEADERS: 3,
        REQUEST_BODY: 4,
        EOF: 5,
    };

    private static LineSeparator = /\r\n/;
    private static HeaderKVSeparator = /:/;
    private static HeaderKVSeparatorChar = ':';
    private static BatchResponse1 = /^--batchresponse_.+--$/i;
    private static BatchResponse2 = /^--batchresponse_.+$/i;
    private static BatchResponse3 = /^HTTP\/1\.1\s+(\d+)\s+(.+)$/i;

    public static Parse(raw: string): ResponseParserPayload[] | any {

        if (!isString(raw))
            return;

        try {
            return JSON.parse(raw);
        } catch (ex) {
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

    public static SafeParse(raw: any): ResponseParserPayload | any {

        if (!isString(raw))
            return raw;

        try {
            return JSON.parse(raw);
        } catch (ex) {
        }

        let dom: DOMParser | undefined;

        try {
            dom = new DOMParser();
        } catch (ex) {
        }

        if (!dom)
            return raw;

        let doc: Document | undefined;

        try {
            doc = dom.parseFromString(raw, 'text/xml');
        } catch (ex) {
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
    private options: BatchJobOptions;

    /** @internal */
    private responsePayload: ResponseParserPayload | string | undefined;

    constructor(options: BatchJobOptions) {
        //@ts-expect-error
        this.options = extend({}, FallbackBatchJobOptions, options);
    }

    public getOptions() {
        return this.options;
    }

    /**
     * Returns the `GET`, `POST`, `...` method.
     * @returns Request method.
     */
    public getMethod(): RequestMethods {
        return this.options.method;
    }

    /**
     * Returns the full request URL.
     * @returns Request URL with any optional params.
     */
    public getUrl() {
        return `${this.options.url}${toParams(this.options.params)}`;
    }

    /**
     * Returns the headers object for the request.
     * @returns Request header array.
     */
    public getHeaders(): BatchJobHeader[] {

        const headers = this.options.headers;
        const results: BatchJobHeader[] = [];

        if (isObject(headers)) {
            for (const key in headers) {
                //@ts-ignore
                const value = headers[key];
                results.push({ key, value });
            }
        }

        return results;

    }

    /** @internal */
    public getPayload(): BodyInit | null | undefined {
        return this.options.body;
    }

    /**
     * Returns the payload from the server once the changeset is queried and processed.
     * @returns The raw data the server responded with from the request.
     */
    public getResponsePayload() {
        return this.responsePayload;
    }

    /** @internal */
    public setResponsePayload(payload: ResponseParserPayload | string | undefined) {
        this.responsePayload = payload;
    }

}

/**
 * Bundles of Changeset objects are lumped into "batch job" entries.
 * @internal
 */
class BatchJob {

    public static NumMaxChangesets = 100;

    /** @internal */
    private options: BatchJobOptions;

    /** @internal */
    private changesets: Changeset[];

    constructor(options: BatchJobOptions) {
        //@ts-expect-error
        this.options = extend({}, FallbackBatchJobOptions, options);
        this.changesets = [];
    }

    public getOptions() {
        return this.options;
    }

    public getChangesets() {
        return this.changesets;
    }

    /** @ignore */
    public isChangesetsFull() {
        return this.changesets.length >= BatchJob.NumMaxChangesets;
    }

    /**
     * Append a changeset to the changeset queue.
     * @param changeset An object of `Changeset`.
     * @returns Successfull additions return the `index` in the queue otherwise `-1`.
     */
    public addChangeset(changeset: Changeset): number {
        if (changeset instanceof Changeset) {
            const options = changeset.getOptions();
            if (options.url[0] === '/')
                options.url = `${this.options.url}${options.url}`;
            return this.changesets.push(changeset) - 1;
        }
        return -1;
    }

    /** @internal */
    public getPayload(guid: string): string {

        const data = [];

        for (const changeset of this.changesets) {

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
    private getSendOptions(batch: SharePointBatch, options?: RequestOptions): RequestOptions {

        const batchOptions = batch.getOptions();

        //@ts-expect-error
        const fallback = extend({}, FallbackRequestOptions, batchOptions, options) as RequestOptions;

        const guid = createGUID();

        fallback.method = 'POST';
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
    private processResponsePayload(changeset: Changeset, payload: ResponseParserPayload | string | undefined) {

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
    public async send(batch: SharePointBatch, options?: RequestOptions): Promise<SharePointBatchJobResponse> {

        const changesets = this.changesets;
        changesets.forEach(changeset => safeCall(changeset.getOptions(), 'before', changeset));

        const safeDone = () => {
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'done', changeset));
        };

        const safeError = () => {
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'fail', changeset));
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'finally', changeset));
        };

        const safeFinally = () => {
            changesets.forEach(changeset => safeCall(changeset.getOptions(), 'finally', changeset));
        };

        const fallback = this.getSendOptions(batch, options);

        //@ts-expect-error
        const backup = extend({}, FallbackBatchJobOptions, fallback) as RequestOptions;

        // store the state from the original response callbacks
        let delayedDone = false;
        let delayedFail = false;
        let delayedArgs: any[] = [];

        // override the handlers with our own on the fallback object
        fallback.before = () => {};
        fallback.done = (...args: any[]) => (delayedDone = true, delayedArgs = args);
        fallback.fail = (...args: any[]) => (delayedFail = true, delayedArgs = args);
        fallback.finally = () => {};

        // perform the query
        const response = await Request(fallback);

        // drop the fallback option reference from the start of the array
        delayedArgs.shift();

        if (!response) {
            if (delayedFail)
                safeCall(backup, 'fail', response, delayedArgs[3] || delayedArgs[1], 0, delayedArgs[3] || delayedArgs[1]); // override the callback arguments to include data we want made aware of in the parent caller (added `delayedArgs`)
            safeError();
            return;
        }

        const payload = await response.text();

        if (!payload || !response.ok) {
            const safePayload = ResponseParser.SafeParse(payload);
            if (delayedFail)
                safeCall(backup, 'fail', response, safePayload, response.status, response.statusText); // override the callback arguments to include data we want made aware of in the parent caller (added `safePayload`)
            safeError();
            return;
        }

        const parsed = ResponseParser.Parse(payload);

        if (!isArray(parsed)) {
            changesets.forEach(changeset => this.processResponsePayload(changeset, payload));
            if (delayedDone)
                safeCall(backup, 'done', response, payload); // override the callback arguments to include data we want made aware of in the parent caller (added `payload`)
            safeFinally();
            return payload;
        }

        const changesetPayloads = parsed as ResponseParserPayload[];
        changesets.forEach((changeset, index) => this.processResponsePayload(changeset, changesetPayloads[index]));
        if (delayedDone)
            safeCall(backup, 'done', response, changesetPayloads); // override the callback arguments to include data we want made aware of in the parent caller (added `changesetPayloads`)
        safeDone();
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
    public static GetSharePointOptions(): SharePointOptions | undefined {

        const win: any = window;
        const context: any = win._spPageContextInfo;
        const getDigest: any = win.GetRequestDigest;

        if (!isObject(context) || typeof getDigest !== 'function')
            return;

        let url: any = context.webAbsoluteUrl;
        let digest: any;

        try {
            digest = getDigest();
        } catch (ex) {
        }

        if (!isString(url))
            url = '';

        if (!isString(digest))
            digest = '';

        return { url, digest };

    }

    /** @internal */
    private options: SharePointOptions;

    /** @internal */
    private jobs: BatchJob[];

    /** @internal */
    private job: BatchJob | undefined;

    constructor(options: SharePointOptions) {
        this.options = options;
        this.jobs = [];
    }

    public getOptions() {
        return this.options;
    }

    /** @internal */
    private appendNewJob(options?: BatchJobOptions): BatchJob {

        //@ts-expect-error
        const fallback = extend({}, FallbackBatchJobOptions, this.options, options) as BatchJobOptions;

        this.job = new BatchJob(fallback);
        this.jobs.push(this.job);

        return this.job;

    }

    /** @internal */
    private getActiveJob(): BatchJob {
        if (!this.job)
            return this.appendNewJob();
        return this.job;
    }

    /**
     * Append a changeset to the batch queue.
     * @param changeset Object instance of `Changeset`.
     * @returns `true` if the changeset was added otherwise `false`.
     */
    public add(changeset: Changeset): boolean {
        let job = this.getActiveJob();
        if (job.isChangesetsFull())
            job = this.appendNewJob();
        for (const j of this.jobs)
            if (j.getChangesets().indexOf(changeset) > -1)
                return false;
        return job.addChangeset(changeset) > -1;
    }

    /**
     * Remove a changeset from the batch queue.
     * @param changeset Object instance of `Changeset`.
     * @returns `true` if the changeset was removed otherwise `false`.
     */
    public remove(changeset: Changeset): boolean {
        for (const job of this.jobs) {
            const changesets = job.getChangesets();
            const i = changesets.indexOf(changeset);
            if (i === -1)
                continue;
            changesets.splice(i, 1);
            return true;
        }
        return false;
    }

    /**
     * Remove all changesets from the batch queue.
     */
    public clear() {
        this.jobs.forEach(job => {
            const changesets = job.getChangesets();
            changesets.splice(0, changesets.length);
        });
        while (this.jobs.length > 1)
            this.jobs.pop();
        this.job = this.jobs[0];
    }

    /**
     * Returns all changesets in the batch queue.
     * @returns Array of changesets.
     */
    public getChangesets(): Changeset[] {
        return this.jobs.reduce((pv, cv) => { pv.push(...cv.getChangesets()); return pv; }, [] as Changeset[]);
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
    public async send(options?: SharePointBatchOptions): Promise<SharePointBatchResponse> {

        //@ts-expect-error
        const fallback = extend({}, FallbackBatchJobOptions, options) as RequestOptions;

        //@ts-expect-error
        const backup = extend({}, FallbackBatchJobOptions, fallback) as SharePointBatchOptions;

        // store the state from the original response callbacks
        let delayedDone = false;
        let delayedFail = false;
        let delayedArgs: any[] = [];

        // override the handlers with our own on the fallback object
        fallback.before = () => {};
        fallback.done = (...args: any[]) => (delayedDone = true, delayedArgs = args);
        fallback.fail = (...args: any[]) => (delayedFail = true, delayedArgs = args);
        fallback.finally = () => {};

        const results: SharePointBatchJobResponse[] = [];
        let success = 0;
        let fail = 0;

        for (const job of this.jobs) {

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
            delayedArgs[1] = backup;
        }

        // drop the fallback option reference from the start of the array
        delayedArgs.shift();

        // if success we return the appropriate data
        if (delayedDone) {
            delayedArgs[1] = results;
            safeCall(backup, 'done', ...delayedArgs);
            safeCall(backup, 'finally', ...delayedArgs);
            return { success: true, ok: true, results } as SharePointBatchResponseSuccess;
        }

        // otherwise return failed result from the response
        safeCall(backup, 'fail', ...delayedArgs);
        safeCall(backup, 'finally', ...delayedArgs);
        return { error: true, ok: false, results: delayedArgs[1] } as SharePointBatchResponseError;

    }

}

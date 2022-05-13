/*!
 * SharePointBatch
 * Copyright 2022 Alex Pedersen
 * Licensed under the MIT license
 * https://github.com/Vladinator89/sharepoint-batch
 */

(() => {

    type BatchOptions = {
        url: string;
        digest: string;
    };

    type BatchJobOptions = {
        guid: string;
        method: 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';
        url: string;
        headers: Record<string, string>;
        changesets: BatchJobOptions[] | null[];
        args: Record<string, string>;
        params: Record<string, string>;
        data: Document | string | Blob | BufferSource | FormData | URLSearchParams | null | undefined;
    };

    type AjaxOptions = {
        xhr: XMLHttpRequest;
        method: 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';
        url: string;
        headers: Record<string, string>;
        mime: string;
        data: Document | string | Blob | BufferSource | FormData | URLSearchParams | null | undefined;
        before?: Function;
        progress?: Function;
        done?: Function;
        fail?: Function;
        always?: Function;
        after?: Function;
    };

    class BatchJob {

        _options: BatchJobOptions;
        _data: string[];
        _changesetResults: object[] | null;

        constructor(options: BatchJobOptions) {

            this._options = options;
            this._data = [];
            this._changesetResults = null;

            if (!SharePointBatch.isArray(options.changesets) || !options.changesets.length)
                options.changesets = [null];

            if (options.method !== 'GET') {
                this._post();
            } else {
                this._get();
            }

        }

        _post() {

            const options = this._options;
            const data = this._data;
            const boundary = `changeset_${SharePointBatch.createGUID()}`;

            data.push(`--batch_${options.guid}`);

            data.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
            data.push('Content-Transfer-Encoding: binary');
            data.push('');

            for (const changeset of options.changesets) {

                data.push(`--${boundary}`);
                data.push('Content-Type: application/http');
                data.push('Content-Transfer-Encoding: binary');
                data.push('');

                data.push(`${options.method} ${options.url + SharePointBatch.toParams(options.params)} HTTP/1.1`);
                data.push('Accept: application/json;odata=verbose');
                data.push('Content-Type: application/json;odata=verbose');

                let headers = options.headers;

                if (changeset) {
                    headers = SharePointBatch.extend({}, headers) as Record<string, string>;
                    SharePointBatch.extend(headers, changeset.headers);
                }

                if (SharePointBatch.isObject(headers, true)) {
                    for (const header in headers) {
                        data.push(`${header}: ${headers[header]}`);
                    }
                }

                data.push('');

                const jsonPayload = changeset && changeset.data
                    ? JSON.stringify(changeset.data)
                    : null;

                if (jsonPayload) {
                    data.push(jsonPayload);
                    data.push('');
                }

            }

            data.push(`--${boundary}--`);

        }

        _get() {

            const options = this._options;
            const data = this._data;

            data.push(`--batch_${options.guid}`);

            data.push('Content-Type: application/http');
            data.push('Content-Transfer-Encoding: binary');
            data.push('');

            for (const changeset of options.changesets) {

                data.push(`${options.method} ${options.url + SharePointBatch.toParams(options.params)} HTTP/1.1`);
                data.push('Accept: application/json;odata=verbose');

                if (SharePointBatch.isObject(options.headers, true)) {

                    for (const header in options.headers) {
                        data.push(`${header}: ${options.headers[header]}`);
                    }

                }

                data.push('');

            }

        }

        payload(): string {
            return this._data.join('\r\n');
        }

        toString(): string {
            return this.payload();
        }

        decorateJob(index: number, results: any[]): number {

            if (!this._changesetResults) {
                this._changesetResults = [];
            } else {
                this._changesetResults.splice(0, this._changesetResults.length);
            }

            for (let i = 0; i < this._options.changesets.length; i++) {

                const result = results[index++];

                if (SharePointBatch.isObject(result, true)) {
                    result.job = this;
                    result.changeset = i;
                }

                this._changesetResults.push(result);

            }

            return index;

        }

    }

    class SharePointBatch {

        static createGUID(): string {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c: string) => {
                const r = Math.random() * 16 | 0;
                return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
            });
        }

        static toParams(object: Object): string {

            const arrayPrefix = 'a';
            const paramPrefix = '?';
            const paramDelim = '&';

            const serialize = (o: object, n?: string): string => {

                if (Array.isArray(o)) {

                    return o.map((v, k) => `${n || arrayPrefix}[${k}]=${serialize(v)}`).join(paramDelim);

                } else if (this.isObject(o)) {

                    const p: string[] = [];

                    for (const k in o) {

                        if (!o.hasOwnProperty(k))
                            continue;

                        const v = o[k];

                        if (Array.isArray(v)) {
                            p.push(serialize(o[k], k));
                        } else {
                            p.push(`${k}=${serialize(o[k])}`);
                        }

                    }

                    return p.join(paramDelim);

                }

                const p = `${o}`;

                try {
                    return encodeURIComponent(p);
                } catch (ex) {
                }

                return p;

            };

            const params = serialize(object);

            if (!params.length)
                return '';

            return `${paramPrefix}${params}`;

        }

        static extend(...objects: Object[]): Object {

            const deep = objects[0] === true;
            const target = deep ? objects[1] : objects[0];
            const sources = deep ? objects.splice(0, 1) : objects;
            return Object.assign(target, ...sources);

        }

        static isObject(object: Object, isPureObject?: boolean): boolean {
            return object && typeof object === 'object' && (!isPureObject || !this.isArray(object));
        }

        static isArray(object: Object): boolean {
            return Array.isArray(object);
        }

        static ajax(options: Partial<AjaxOptions>) {

            const xhr = new XMLHttpRequest();

            const fallback: Partial<AjaxOptions> = {
                xhr: xhr,
                method: 'GET',
                headers: {},
            };

            this.extend(options, fallback);

            const progress = (_: any, ...args: any) => {
                if (options.progress)
                    options.progress.call(options, options, ...args);
            };

            const load = (_: any, ...args: any) => {
                if ((xhr.status / 100 | 0) !== 2)
                    return error.call(options, _, ...args);
                if (options.done)
                    options.done.call(options, options, ...args);
                if (options.always)
                    options.always.call(options, options, ...args);
                if (options.after)
                    options.after.call(options, options, ...args);
            };

            const error = (_: any, ...args: any) => {
                if (options.fail)
                    options.fail.call(options, options, ...args);
                if (options.always)
                    options.always.call(options, options, ...args);
                if (options.after)
                    options.after.call(options, options, ...args);
            };

            xhr.addEventListener('progress', progress);
            xhr.addEventListener('load', load);
            xhr.addEventListener('error', error);
            xhr.addEventListener('abort', error);

            if (options.method && options.url)
                xhr.open(options.method, options.url);

            xhr.setRequestHeader('Accept', 'application/json;odata=verbose');

            if (options.headers && this.isObject(options.headers, true)) {
                for (const header in options.headers) {
                    xhr.setRequestHeader(header, options.headers[header]);
                }
            }

            if (options.mime) {
                xhr.overrideMimeType(options.mime);
            }

            if (options.before)
                options.before.call(options, options);

            xhr.send(options.data);

            return options;

        }

        _options: BatchOptions;
        _guid: string;
        _jobs: BatchJob[];
        _changesetSlots: number;

        constructor(options: BatchOptions) {
            this._options = SharePointBatch.extend({}, options) as BatchOptions;
            this._guid = SharePointBatch.createGUID();
            this._jobs = [];
            this._changesetSlots = 100;
        }

        append(options: BatchJobOptions): number {
            let changesetSlots = this._changesetSlots;
            changesetSlots -= options.changesets.length;
            if (changesetSlots < 0)
                return -1;
            this._changesetSlots = changesetSlots;
            const job = new BatchJob(options);
            return this._jobs.push(job) - 1;
        }

        remove(index: number): BatchJob | null {
            return this._jobs.splice(index, 1)[0];
        }

        payload(): string {
            return this._jobs.map(job => job.payload()).join('\r\n');
        }

        toString(): string {
            return this.payload();
        }

        send(options: AjaxOptions) {

            const fallback: Partial<AjaxOptions> = {
                method: 'POST',
                url: `${this._options.url}/_api/$batch`,
                headers: {
                    'X-RequestDigest': this._options.digest,
                    'Content-Type': `multipart/mixed; boundary="batch_${this._guid}"`,
                },
                data: `${this.payload()}--batch_${this._guid}--`,
            };

            SharePointBatch.extend(options, fallback);

            const backup = SharePointBatch.extend({}, options) as AjaxOptions;

            options.done = () => {
                const json = this._parseResponse(options.xhr.responseText);
                this._decorateJobs(json);
                if (backup.done)
                    backup.done.call(backup, backup, json);
            };

            options.fail = () => {
                const json = this._parseResponse(options.xhr.responseText);
                this._decorateJobs(json);
                if (backup.fail)
                    backup.fail.call(backup, backup, json);
            };

            options.always = () => {
                if (backup.always)
                    backup.always.call(backup, backup);
            };

            return SharePointBatch.ajax(options);

        }

        spawn(options: BatchOptions) {
            return new SharePointBatch(options);
        }

        static _parseLevels: Record<string, number> = {
            UNKNOWN: 0,
            HEADERS: 1,
            REQUEST: 2,
            REQUEST_HEADERS: 3,
            REQUEST_BODY: 4,
            EOF: 5,
        };

        static _parseLineSeparatorPattern = /\r\n/;
        static _parseHeaderSeparatorPattern = /:/;
        static _parseHeaderSeparator = ':';
        static _parseBatchResponsePattern1 = /^--batchresponse_.+--$/i;
        static _parseBatchResponsePattern2 = /^--batchresponse_.+$/i;
        static _parseBatchResponsePattern3 = /^HTTP\/1\.1\s+(\d+)\s+(.+)$/i;

        _parseResponse(raw?: string): any {

            if (!raw)
                return null;

            try {
                return JSON.stringify(raw);
            } catch (ex) {
            }

            const parseLevels = SharePointBatch._parseLevels;
            const lines = raw.split(SharePointBatch._parseLineSeparatorPattern);
            const results: string[] = [];

            let temp: any = null;
            let cwo;
            let level = parseLevels.UNKNOWN;

            for (const line of lines) {

                if (SharePointBatch._parseBatchResponsePattern1.test(line)) {

                    if (temp) {
                        temp.data = this._parseResponse(temp.data);
                        results.push(temp);
                        temp = null;
                    }

                    level = parseLevels.EOF;
                    break;

                } else if (SharePointBatch._parseBatchResponsePattern2.test(line)) {

                    if (temp) {
                        temp.data = this._parseResponse(temp.data);
                        results.push(temp);
                    }

                    temp = {
                        headers: {},
                        http: { status: 0, statusText: '' },
                        data: null,
                    };

                    cwo = temp;
                    level = parseLevels.HEADERS;

                } else if (level === parseLevels.REQUEST_BODY) {

                    if (cwo.data == null)
                        cwo.data = '';

                    cwo.data += line;

                } else if (SharePointBatch._parseBatchResponsePattern3.test(line)) {

                    if (level === parseLevels.REQUEST) {

                        const http = line.match(SharePointBatch._parseBatchResponsePattern3);
                        cwo.http.status = http && parseInt(http[1], 10);
                        cwo.http.statusText = http && http[2];
                        level = parseLevels.REQUEST_HEADERS;

                    }

                } else if (/^.+:\s*.+$/i.test(line)) {

                    if (level === parseLevels.HEADERS || level === parseLevels.REQUEST_HEADERS) {

                        const parts = line.split(SharePointBatch._parseHeaderSeparatorPattern);

                        if (parts) {

                            const key = parts.shift()?.trim();

                            if (key)
                                cwo.headers[key] = parts.join(SharePointBatch._parseHeaderSeparator).trim();

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

            return results;

        }

        _decorateJobs(results: object[]) {
            let index = 0;
            for (const job of this._jobs)
                index = job.decorateJob(index, results);
        }

    }

    class SharePointBatchUtil extends SharePointBatch {

        __options: BatchOptions & BatchJobOptions;
        __jobs: SharePointBatch[];
        __currentJob: SharePointBatch;

        constructor(options: BatchOptions & BatchJobOptions) {
            super(options);
            this.__options = options;
            this.__jobs = [];
            this.__currentJob = new SharePointBatch(options);
            this.__jobs.push(this.__currentJob);
        }

        _spawn(): SharePointBatchUtil {
            return new SharePointBatchUtil(this.__options);
        }

        _append(...args: any): number {

            let count = -1;

            for (let i = 0; i < args.length; i++) {

                const arg = args[i];

                if (arg instanceof SharePointBatch) {

                    this.__jobs.push(arg);
                    count++;

                } else if (SharePointBatch.isObject(arg, true)) {

                    let index = this.__currentJob.append(arg);
                    count++;

                    if (index < 0) {
                        this.__currentJob = this.__currentJob.spawn(this.__currentJob._options);
                        this.__currentJob.append(arg);
                        this.__jobs.push(this.__currentJob);
                    }

                }

            }

            return count;

        }

        _remove(...args: any): object[] | null {

            const purged: any[] = [];

            for (let i = 0; i < args.length; i++) {

                const arg = args[i];

                let index = this.__jobs.indexOf(arg);
                if (index > -1)
                    purged.push(...this.__jobs.splice(index, 1));

                for (let j = 0; j < this.__jobs.length; j++)  {

                    const job = this.__jobs[j];

                    if (!SharePointBatch.isObject(job._jobs, true))
                        continue;

                    index = job._jobs.indexOf(arg);
                    if (index > -1)
                        purged.push(...job._jobs.splice(index, 1));

                }

            }

            return purged.length
                ? purged
                : null;

        }

        _send(options: BatchOptions & BatchJobOptions & AjaxOptions) {

            options = SharePointBatch.extend({}, options) as BatchOptions & BatchJobOptions & AjaxOptions;

            const jobs = this.__jobs;
            const jobResults: object[] = [];
            let jobIndex = 0;

            if (options.before)
                options.before.call(options, options, jobs, jobResults);

            const next = () => {

                const job = jobs[jobIndex];

                if (job) {

                    const done = (_options: any, data: any) => {
                        mergeResults(data);
                    };

                    const fail = (_options: any, data: any) => {
                        mergeResults(data);
                    };

                    const always = (_options: any, data: any) => {
                        if (options.progress)
                            options.progress.call(options, options, job, jobResults);
                        jobIndex++;
                        next();
                    };

                    const mergeResults = (data: object[]) => {
                        if (SharePointBatch.isArray(data)) {
                            jobResults.push(...data);
                        } else {
                            jobResults.push(data);
                        }
                    };

                    const modified = SharePointBatch.extend({}, options, {
                        done,
                        fail,
                        always,
                    } as Partial<AjaxOptions>) as AjaxOptions;

                    job.send(modified);

                } else {

                    if (options.done)
                        options.done.call(options, options, jobs, jobResults);

                    if (options.always)
                        options.always.call(options, options, jobs, jobResults);

                    if (options.after)
                        options.after.call(options, options, jobs, jobResults);

                }

            };

            next();

        }

    }

    (window as any).SharePointBatch = SharePointBatch;
    (window as any).SharePointBatchUtil = SharePointBatchUtil;

})();

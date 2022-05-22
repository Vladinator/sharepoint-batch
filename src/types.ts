import { Changeset } from './sharepoint';

/**
 * Callback options.
 */
export type Callbacks<B, D, F, A> = {
    before?: B;
    done?: D;
    fail?: F;
    finally?: A;
}

/**
 * Callback properties.
 */
export type CallbackProps = 'before' | 'done' | 'fail' | 'finally';

/**
 * Http request methods.
 */
export type RequestMethods = 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';

export type RequestCallback = (options: RequestOptions, response: RequestResponse) => void;
export type RequestCallbackFail = (options: RequestOptions, response: RequestResponse, status?: number | string, statusText?: string) => void;

export type RequestOptionsBase = RequestInit & {
    method: RequestMethods;
    url: string;
};

/**
 * Request options with callback options.
 */
export type RequestOptions = RequestOptionsBase & Callbacks<RequestCallback, RequestCallback, RequestCallbackFail, RequestCallback>;

/**
 * Request response.
 */
export type RequestResponse = Response | undefined;

/**
 * The SharePoint hostWeb url and security digest.
 */
export type SharePointOptions = {
    url: string;
    digest: string;
};

/**
 * Http params record.
 */
export type SharePointParams = Record<string, any> | undefined;

export type BatchJobCallback = (options: BatchJobOptions, changeset: Changeset) => void;
export type BatchJobCallbackResponse = (options: BatchJobOptions, changeset: Changeset, response: ResponseParserPayload) => void;

/**
 * Request options with optional params array.
 */
export type BatchJobOptions = RequestOptionsBase & Callbacks<BatchJobCallback, BatchJobCallbackResponse, BatchJobCallbackResponse, BatchJobCallback> & {
    params?: SharePointParams;
};

/**
 * A batch job header record.
 */
export type BatchJobHeader = {
    key: string;
    value: string;
};

/**
 * A batch job response is either an array of payload records, a string or nothing.
 */
export type SharePointBatchJobResponse = ResponseParserPayload[] | string | undefined;

/**
 * A batch response is either the success or error object.
 */
export type SharePointBatchResponse = SharePointBatchResponseSuccess | SharePointBatchResponseError;

/**
 * A changeset response success.
 */
export type SharePointBatchResponseSuccess = {
    success: true;
    ok: boolean;
    results: SharePointBatchJobResponse[];
};

/**
 * A changeset response error.
 */
export type SharePointBatchResponseError = {
    error: true;
    ok: boolean;
    results: any;
};

/**
 * Request response header record.
 */
export type ResponseParserHeaders = Record<string, string>;

/**
 * Request response http status.
 */
export type ResponseParserHttp =  {
    status: number;
    statusText: string;
};

/**
 * A changeset response payload.
 */
export type ResponseParserPayload = {
    changeset?: Changeset;
    headers: ResponseParserHeaders;
    http: ResponseParserHttp;
    ok: boolean;
    data: any;
};

export type SharePointBatchCallback = (options: SharePointBatchOptions, response: RequestResponse) => void;
export type SharePointBatchCallbackResult = (options: SharePointBatchOptions, response: RequestResponse, result: SharePointBatchResponse) => void;
export type SharePointBatchCallbackFail = (options: SharePointBatchOptions, response: RequestResponse, result: SharePointBatchResponse | string, status: number, statusText: string) => void;

/**
 * Special callback parameters are used in the batch callback options.
 */
export type SharePointBatchOptions = RequestOptionsBase & Callbacks<SharePointBatchCallback, SharePointBatchCallbackResult, SharePointBatchCallbackFail, SharePointBatchCallbackResult>;

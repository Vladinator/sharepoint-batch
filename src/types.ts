import { Changeset } from './sharepoint';

/**
 * Callback options.
 */
export type CallbackOptions = {
    before?: Function;
    done?: Function;
    fail?: Function;
    finally?: Function;
}

/**
 * Callback properties.
 */
export type CallbackProps = 'before' | 'done' | 'fail' | 'finally';

/**
 * Http request methods.
 */
export type RequestMethods = 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';

/**
 * Request options with callback options.
 */
export type RequestOptions = CallbackOptions & RequestInit & {
    method: RequestMethods;
    url: string;
};

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

/**
 * Request options with optional params array.
 */
export type BatchJobOptions = RequestOptions & {
    params?: SharePointParams[];
};

/**
 * A batch job header record.
 */
export type BatchJobHeader = {
    key: string;
    value: any;
};

/**
 * A batch response is either an array of payload records, a string or nothing.
 */
export type SharePointBatchResponse = ResponseParserPayload[] | string | undefined;

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

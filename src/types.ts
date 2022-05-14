export type CallbackOptions = {
    before?: Function;
    done?: Function;
    fail?: Function;
    finally?: Function;
}

export type CallbackProps = 'before' | 'done' | 'fail' | 'finally';

export type RequestOptions = CallbackOptions & RequestInit & {
    method: 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';
    url: string;
}

export type RequestResponse = Response | undefined;

export type RequestReturnType = JSON | Object | String | Blob | ArrayBuffer | FormData | RequestResponse;

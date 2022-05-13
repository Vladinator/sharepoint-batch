export type CallbackOptions = {
    before?: Function;
    progress?: Function;
    done?: Function;
    fail?: Function;
    always?: Function;
    after?: Function;
}

export type RequestOptions = CallbackOptions & RequestInit & {
    method: 'GET' | 'HEAD' | 'POST' | 'PUT' | 'DELETE' | 'CONNECT' | 'OPTIONS' | 'TRACE' | 'PATCH';
    url: string;
}

import { CallbackProps, RequestOptions, RequestResponse } from './types';

const SafeCall = (options: RequestOptions, prop: CallbackProps, response: RequestResponse, ...args: any) => {
    const value = options[prop];
    if (typeof value === 'function')
        value.call(null, options, response, ...args);
};

export const Request = async (options: RequestOptions): Promise<RequestResponse> => {

    return new Promise(async resolve => {

        let response: RequestResponse;
        SafeCall(options, 'before', response);

        try {
            response = await fetch(options.url, options as never);
            SafeCall(options, 'done', response);
        } catch (ex: any) {
            SafeCall(options, 'fail', response, ex);
        }

        SafeCall(options, 'finally', response);
        resolve(response);

    });

};

type ObjectToMethodMapRecord = {
    type: any;
    prop: 'arrayBuffer' | 'blob' | 'formData' | 'json' | 'text';
};

const ObjectToMethodMap: ObjectToMethodMapRecord[] = [
    { type: ArrayBuffer, prop: 'arrayBuffer' },
    { type: Blob, prop: 'blob' },
    { type: FormData, prop: 'formData' },
    { type: Object, prop: 'json' },
    { type: String, prop: 'text' },
];

export const RequestResolve = async <T>(options: RequestOptions, type?: T): Promise<T | undefined> => {

    const response = await Request(options);

    if (!response)
        return;

    if (!type || type instanceof Response)
        return response as never;

    for (const map of ObjectToMethodMap) {
        if (type instanceof map.type) {
            try {
                const result = await response[map.prop]();
                return result as T;
            } catch (ex: any) {
                return;
            }
        }
    }

};

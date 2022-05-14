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

export const RequestArrayBuffer = async (options: RequestOptions): Promise<ArrayBuffer | undefined> => {
    const response = await Request(options);
    if (!response)
        return;
    try {
        return await response.arrayBuffer();
    } catch (ex) {
    }
};

export const RequestBlob = async (options: RequestOptions): Promise<Blob | undefined> => {
    const response = await Request(options);
    if (!response)
        return;
    try {
        return await response.blob();
    } catch (ex) {
    }
};

export const RequestFormData = async (options: RequestOptions): Promise<FormData | undefined> => {
    const response = await Request(options);
    if (!response)
        return;
    try {
        return await response.formData();
    } catch (ex) {
    }
};

export const RequestJson = async (options: RequestOptions): Promise<any | undefined> => {
    const response = await Request(options);
    if (!response)
        return;
    try {
        return await response.json();
    } catch (ex) {
    }
};

export const RequestText = async (options: RequestOptions): Promise<string | undefined> => {
    const response = await Request(options);
    if (!response)
        return;
    try {
        return await response.text();
    } catch (ex) {
    }
};

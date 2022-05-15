import { RequestOptions, RequestResponse } from './types';
import { safeCall } from './utils';

export const Request = async (options: RequestOptions): Promise<RequestResponse> => {

    return new Promise(async resolve => {

        let response: RequestResponse;
        safeCall(options, 'before', response);

        try {
            response = await fetch(options.url, options as never);
            if (response && response.ok) {
                safeCall(options, 'done', response);
            } else {
                safeCall(options, 'fail', response, response?.status, response?.statusText);
            }
        } catch (ex: any) {
            safeCall(options, 'fail', response, ex);
        }

        safeCall(options, 'finally', response);
        resolve(response);

    });

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

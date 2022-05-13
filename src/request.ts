import { RequestOptions } from './types';

const SafeCall = (options: RequestOptions, prop: string, response: Response | undefined, ...args: any) => {
    const value = options[prop] as Function;
    if (typeof value === 'function')
        value.call(options, options, response, ...args);
};

export const Request = async (options: RequestOptions): Promise<Response | undefined> => {

    return new Promise(async resolve => {

        let response: Response | undefined;

        SafeCall(options, 'before', response);

        try {
            response = await fetch(options.url, options as never);
            SafeCall(options, 'progress', response);
            SafeCall(options, 'done', response);
        } catch (ex) {
            console.error(ex);
            SafeCall(options, 'fail', response);
        }

        SafeCall(options, 'always', response);
        SafeCall(options, 'after', response);

        resolve(response);

    });

};

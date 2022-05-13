import { RequestOptions } from 'types';
import { extend } from 'utils';

const fallback: Partial<RequestOptions> = { method: 'GET' };

export const Request = async (options: RequestOptions) => {

    extend(options, fallback);

    const response = await fetch(options.url, options as never);

    console.warn('Request', options, '->', response.ok, response.status, response.statusText); // DEBUG

    return response;

};

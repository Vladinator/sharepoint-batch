import fetch from 'node-fetch';
import { extend } from 'utils';
import { RequestOptions } from 'types';

const fallback: Partial<RequestOptions> = { method: 'GET' };

export const Request = async (options: RequestOptions) => {

    extend(options, fallback);

    const response = await fetch(options.url, options as never);

    console.warn(response.ok, response.status, response.statusText); // DEBUG

    return response;

};

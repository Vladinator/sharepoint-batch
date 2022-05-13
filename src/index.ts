import { RequestOptions } from 'https';
import { Request } from './request';

function debug(name: string) {
    return function(options: RequestOptions, response: Response) {
        console.warn('debug', name, '|', options, '->', response); // DEBUG
    }
}

(async () => {

    const responses = [
        await Request({
            method: 'GET',
            url: 'https://127.0.0.1',
            before: debug('before'),
            progress: debug('progress'),
            done: debug('done'),
            fail: debug('fail'),
            always: debug('always'),
            after: debug('after'),
        }),
        await Request({
            method: 'GET',
            url: 'http://127.0.0.1',
            before: debug('before'),
            progress: debug('progress'),
            done: debug('done'),
            fail: debug('fail'),
            always: debug('always'),
            after: debug('after'),
        }),
    ];

    for (const response of responses) {
        console.warn(response, '->', response?.ok, response?.status, response?.statusText);
    }

})();

import { equal } from 'assert';
import { Request, RequestJson } from '../src/request';
import { RequestOptions } from '../src/types';

describe('request', () => {

    it('localhost invalid url', () => {
        const promise = Request({ method: 'GET', url: 'http://127.0.0.1:999999' } as RequestOptions);
        promise.then(value => {
            equal(value, undefined, 'malformed request expects nothing');
        }, error => {
            equal(true, false, `malformed request should not yield error: ${error}`);
        });
    });

    it('localhost json invalid url', () => {
        const promise = RequestJson({ method: 'GET', url: 'http://127.0.0.1:999999' } as RequestOptions);
        promise.then(value => {
            equal(value, undefined, 'malformed request expects nothing');
        }, error => {
            equal(true, false, `malformed request should not yield error: ${error}`);
        });
    });

});

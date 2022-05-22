import { equal, deepEqual } from 'assert';
import { BatchJobHeader, BatchJobOptions } from '../src/types';
import { Changeset, SharePointBatch } from '../src/sharepoint';
import { isObject, toParams } from '../src/utils';

describe('sharepoint', () => {

    const convertToKVPArray = (headers?: HeadersInit): BatchJobHeader[] => {
        const temp: BatchJobHeader[] = [];
        if (isObject(headers)) {
            for (const key in headers) {
                temp.push({ key, value: headers[key] });
            }
        }
        return temp;
    };

    const simpleTest = (changeset: Changeset, options: BatchJobOptions): void => {
        deepEqual(changeset.getOptions(), options, 'both options passed and retrieved must be equal');
        equal(changeset.getMethod(), options.method);
        equal(changeset.getUrl(), `${options.url}${toParams(options.params)}`);
        deepEqual(changeset.getHeaders(), convertToKVPArray(options.headers));
        equal(changeset.getPayload(), options.body);
    };

    it('changeset simple', () => {
        const options: BatchJobOptions = { method: 'GET', url: '' };
        const changeset = new Changeset(options);
        simpleTest(changeset, options);
    });

    it('changeset advanced', () => {
        const options: BatchJobOptions = {
            method: 'POST',
            url: 'http://localhost:1234',
            body: JSON.stringify({ hello: 'world' }),
            headers: { 'Key': 'Value' },
            params: { 'hello': 'world' }
        };
        const changeset = new Changeset(options);
        simpleTest(changeset, options);
        let testResponsePayload: any = undefined;
        equal(changeset.getResponsePayload(), testResponsePayload);
        testResponsePayload = 'hello world';
        changeset.setResponsePayload(testResponsePayload);
        equal(changeset.getResponsePayload(), testResponsePayload);
        testResponsePayload = { hello: 'world' };
        changeset.setResponsePayload(testResponsePayload);
        equal(changeset.getResponsePayload(), testResponsePayload);
    });

});

import { equal, deepEqual, notDeepEqual } from 'assert';
import { RequestOptions } from '../src/types';
import { createGUID, extend, isArray, isObject, isString, safeCall, toParams } from '../src/utils';

describe('util', () => {

    it('createGUID', () => {
        const guid = createGUID();
        const valid = /^[a-f0-9]{8}-[a-f0-9]{4}-4[a-f0-9]{3}-[a-f0-9]{4}-[a-f0-9]{12}$/.test(guid);
        equal(valid, true);
    });

    it('extend simple', () => {
        const destination = {};
        const source1 = { a: 1 };
        const source2 = { a: 8, b: 2 };
        const source3 = { a: 9, c: 3 };
        const combined = extend(destination, source1, source2, source3);
        equal(destination === combined, true, 'destination and combined should be the same object');
        equal(destination === source1 || destination === source2 || destination === source3, false, 'destination should not be any of the source object');
        const expects = { a: 9, b: 2, c: 3 };
        deepEqual(destination, expects, 'destination should contain these exact values');
    });

    it('extend deep', () => {
        const destination = {};
        const source1 = {
            a: 9,
            x: { x: 'x' },
            z: { z: { z: 'z' }, zz: 'zz' }
        };
        const source2 = {
            a: 1,
            x: { x: { x: 'x' }, xxx: 'xxx' },
            z: { z: 'z', zzz: 'zzz' },
            y: { y: 'y' },
            b: 2
        };
        extend(destination, source1, source2);
        const expects = {
            a: 1,
            b: 2,
            x: { x: { x: 'x' }, xxx: 'xxx' },
            y: { y: 'y' },
            z: { z: 'z', zzz: 'zzz' }
        };
        deepEqual(destination, expects, 'destination should contain these exact values');
        const expectsDifferent: any = {};
        extend(expectsDifferent, expects);
        expectsDifferent.z.zz = 'zz';
        notDeepEqual(destination, expectsDifferent, 'destination should not contain this because of shallow copy');
    });

    const isTestValues = [
        /* 0 */ undefined,
        /* 1 */ null,
        /* 2 */ 1234,
        /* 3 */ 'text',
        /* 4 */ true,
        /* 5 */ false,
        /* 6 */ {},
        /* 7 */ { length: 0 },
        /* 8 */ [],
    ];

    const isTestExpects = {
        isArray: {
            func: isArray,
            expects: [
                /* 0 */ false,
                /* 1 */ false,
                /* 2 */ false,
                /* 3 */ false,
                /* 4 */ false,
                /* 5 */ false,
                /* 6 */ false,
                /* 7 */ false,
                /* 8 */ true,
            ]
        },
        isObject: {
            func: isObject,
            expects: [
                /* 0 */ false,
                /* 1 */ false,
                /* 2 */ false,
                /* 3 */ false,
                /* 4 */ false,
                /* 5 */ false,
                /* 6 */ true,
                /* 7 */ true,
                /* 8 */ true,
            ]
        },
        isString: {
            func: isString,
            expects: [
                /* 0 */ false,
                /* 1 */ false,
                /* 2 */ false,
                /* 3 */ true,
                /* 4 */ false,
                /* 5 */ false,
                /* 6 */ false,
                /* 7 */ false,
                /* 8 */ false,
            ]
        },
        toParams: {
            func: toParams,
            expects: [
                /* 0 */ '',
                /* 1 */ '',
                /* 2 */ '',
                /* 3 */ '',
                /* 4 */ '',
                /* 5 */ '',
                /* 6 */ '',
                /* 7 */ '?length=0',
                /* 8 */ '',
            ]
        },
    };

    const runTestKVPairs = (key: string) => {
        const testExpects = isTestExpects[key];
        testExpects.expects.forEach((expects: any[], index: number) => equal(testExpects.func(isTestValues[index]), expects, `${key} tested ${isTestValues[index]} and expected ${expects}`));
    };

    it('isArray', () => {
        runTestKVPairs('isArray');
    });

    it('isObject', () => {
        runTestKVPairs('isObject');
    });

    it('isString', () => {
        runTestKVPairs('isString');
    });

    it('safeCall', () => {
        let beforeArgs: any[] = [];
        let finallyArgs: any[] = [];
        const options: RequestOptions = {
            method: 'GET',
            url: '',
            before: (...args) => beforeArgs = args,
            finally: (...args) => finallyArgs = args,
        };
        safeCall(options, 'before', 1, 2, 3);
        safeCall(options, 'finally', 'x', 'y', 'z');
        const expectedBeforeArgs = [ options, 1, 2, 3 ];
        const expectedFinallyArgs = [ options, 'x', 'y', 'z' ];
        deepEqual(beforeArgs, expectedBeforeArgs, 'the callback arguments should match these values');
        deepEqual(finallyArgs, expectedFinallyArgs, 'the callback arguments should match these values');
    });

    it('toParams', () => {
        runTestKVPairs('toParams');
        equal(toParams([ 1, 2, 3 ]), '?a%5B%5D=1&a%5B%5D=2&a%5B%5D=3'); // ?a[]=1&a[]=2&a[]=3
        equal(toParams([ 'a', 'b', 'c' ]), '?a%5B%5D=a&a%5B%5D=b&a%5B%5D=c'); // ?a[]=a&a[]=b&a[]=c
        equal(toParams([ 'a', ['x'], 'c' ]), '?a%5B%5D=a&a%5B1%5D%5B%5D=x&a%5B%5D=c'); // ?a[]=a&a[1][]=x&a[]=c
        equal(toParams([ 'a', { x: 'y' }, 'c' ]), '?a%5B%5D=a&a%5B1%5D%5Bx%5D=y&a%5B%5D=c'); // ?a[]=a&a[1][x]=y&a[]=c
        equal(toParams({ '$select': 'Id,Title', '$filter': 'Title ne null', '$expand': 'Author/Title' }), '?%24select=Id%2CTitle&%24filter=Title%20ne%20null&%24expand=Author%2FTitle'); // ?$select=Id,Title&$filter=Title ne null&$expand=Author/Title
        equal(toParams({ 'key': 1234, 'obj': { a: 1, b: 2, c: 3 }, 'arr': [ 'hello', 'world' ] }), '?key=1234&obj%5Ba%5D=1&obj%5Bb%5D=2&obj%5Bc%5D=3&arr%5B%5D=hello&arr%5B%5D=world'); // ?key=1234&obj[a]=1&obj[b]=2&obj[c]=3&arr[]=hello&arr[]=world
    });

});

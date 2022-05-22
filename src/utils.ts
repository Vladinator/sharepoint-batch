import { CallbackProps, RequestOptions, BatchJobOptions, SharePointBatchOptions } from './types';

export const isArray = (object: any) => Array.isArray(object);

export const isObject = (object: any, plainObject: boolean = false): boolean => !!(object && typeof object === 'object' && (!plainObject || !isArray(object)));

export const isString = (object: any) => typeof object === 'string';

export const extend = <T>(target: T, ...sources: T[]): T => Object.assign(target as Object, ...sources);

export const toParams = (object: any, traditional: boolean = false): string => {

    if (!isObject(object))
        return '';

    const bracket = /\[\]$/;
    const results: string[] = [];

    const append = (key: any, val: any): void => {
        val = typeof val === 'function' ? val() : val;
        val = val == null ? '' : val;
        results[results.length] = `${encodeURIComponent(key)}=${encodeURIComponent(val)}`;
    };

    const serialize = (prefix: string, obj: any, root: boolean = false): void => {

        if (isArray(obj)) {

            for (let i = 0; i < obj.length; i++) {
                const val = obj[i];
                if (traditional || bracket.test(prefix)) {
                    append(prefix, val);
                } else {
                    const k = isObject(val) ? i : '';
                    serialize(root ? `a[${k}]` : `${prefix}[${k}]`, val);
                }
            }

        } else if (!traditional && isObject(obj)) {

            for (const key in obj) {
                serialize(root ? `${prefix}${key}` : `${prefix}[${key}]`, obj[key]);
            }

        } else {

            append(prefix, obj);

        }

    };

    if (object == null)
        return '';

    serialize('', object, true);

    const result = results.join('&');

    if (result === '')
        return '';

    return `?${result}`;

};

export const createGUID = (): string => {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c: string) => {
        const r = Math.random() * 16 | 0;
        return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
    });
};

export const safeCall = (options: RequestOptions | BatchJobOptions | SharePointBatchOptions, prop: CallbackProps, ...args: any[]) => {
    const value = options[prop];
    if (typeof value !== 'function')
        return;
    //@ts-expect-error
    value.call(null, options, ...args);
};

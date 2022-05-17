import { CallbackProps, RequestOptions, BatchJobOptions, SharePointBatchOptions } from './types';

export const isArray = (object: any) => Array.isArray(object);

export const isObject = (object: any, plainObject: boolean = false): boolean => !!(object && typeof object === 'object' && (!plainObject || !isArray(object)));

export const isString = (object: any) => typeof object === 'string';

export const extend = <T>(target: T, ...sources: T[]): T => Object.assign(target as Object, ...sources);

export const toParams = (object: any): string => {

    if (object == null)
        return '';

    const arrayPrefix = 'a';
    const paramPrefix = '?';
    const paramDelim = '&';

    const serialize = (o: object, n?: string): string => {

        if (isArray(o)) {

            //@ts-expect-error
            return o.map((v: any, k: any) => `${n || arrayPrefix}[${k}]=${serialize(v)}`).join(paramDelim);

        } else if (isObject(o)) {

            const p: string[] = [];

            for (const k in o) {

                if (!o.hasOwnProperty(k))
                    continue;

                //@ts-ignore
                const v = o[k];

                if (Array.isArray(v)) {
                    p.push(serialize(v, k));
                } else {
                    p.push(`${k}=${serialize(v)}`);
                }

            }

            return p.join(paramDelim);

        }

        const p = `${o}`;

        try {
            return encodeURIComponent(p);
        } catch (ex) {
        }

        return p;

    };

    const params = serialize(object);

    if (!params.length)
        return '';

    return `${paramPrefix}${params}`;

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

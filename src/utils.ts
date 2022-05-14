export const isArray = (object: any) => Array.isArray(object);

export const isObject = (object: any, plainObject: boolean = false) => object && typeof object === 'object' && (!plainObject || !isArray(object));

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
            return o.map((v, k) => `${n || arrayPrefix}[${k}]=${serialize(v)}`).join(paramDelim);

        } else if (isObject(o)) {

            const p: string[] = [];

            for (const k in o) {

                if (!o.hasOwnProperty(k))
                    continue;

                const v = o[k];

                if (Array.isArray(v)) {
                    p.push(serialize(o[k], k));
                } else {
                    p.push(`${k}=${serialize(o[k])}`);
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

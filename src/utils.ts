export const isArray = (object: any) => Array.isArray(object);

export const isObject = (object: any, plainObject: boolean = false) => object && typeof object === 'object' && (!plainObject || !isArray(object));

export const extend = <T>(target: T, ...sources: T[]): T => Object.assign(target as Object, ...sources);

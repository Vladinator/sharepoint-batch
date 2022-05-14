import { RequestOptions } from './types';
import { RequestResolve } from './request';

(async () => {

    const options: RequestOptions = { method: 'GET', url: window.location.href };

    const requests = [
        await RequestResolve<Response>(options),
        await RequestResolve<ArrayBuffer>(options),
        await RequestResolve<Blob>(options),
        await RequestResolve<FormData>(options),
        await RequestResolve<Object>(options),
        await RequestResolve<String>(options),
    ];

    for (let i = 0; i < requests.length; i++) {
        const request = requests[i];
        console.warn(typeof request, request);
    }

})();

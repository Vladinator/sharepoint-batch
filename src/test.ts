import { Request } from './request';

(async () => {

    const response1 = await Request({ method: 'GET', url: 'https://google.com' });
    const response2 = await Request({ method: 'GET', url: 'http://google.com' });

    console.log('response1', response1); // DEBUG
    console.log('response2', response2); // DEBUG

})();

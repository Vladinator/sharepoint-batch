import { Request } from 'request';

(async () => {

    const response1 = await Request({ method: 'GET', url: 'https://google.com' });

    const response2 = await Request({ method: 'GET', url: 'http://google.com' });

})();

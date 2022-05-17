# SharePointBatch
Inspiration drawn from the MSDN article ["Make batch requests with the REST APIs"](https://msdn.microsoft.com/en-us/library/office/dn903506.aspx)

This Javascript library is supposed to help developers efficiently build batch jobs and parsing the response from the server. The reason I made this was that the existing code provided by Microsoft seemed a bit underwhelming and the examples were too specific. There is a need for a general purpose library to utilize the API in existing projects.

Please note that this library is a work in progress. It's a mix between a research project and experimenting writing typescript to compile into JavaScript to then be used in other projects.

## API
The project can be loaded as both a module and a javascript browser script.

### TypeScript
You can import the package as a module. You will need to specify the url and digest manually.

```typescript
import { SharePointBatch, Changeset } from 'sharepoint-batch';
const batch = new SharePointBatch({ url: 'https://my.sharepoint.com', digest: '...' });
batch.add(new Changeset({ method: 'POST', url: '/_api/ContextInfo' }));
batch.add(new Changeset({ method: 'GET', url: '/_api/Site', params: { '$select': 'Id, Url, ReadOnly, WriteLocked' } }));
batch.add(new Changeset({ method: 'GET', url: '/_api/Web', params: { '$select': 'Id, Title, WebTemplate, Created' } }));
const response = await batch.send();
console.log(response.ok ? 'Done!' : 'Fail!', response.results);
```

### JavaScript
The pre-built `build.min.js` file can be loaded directly into a ES6 compatible browser.

```javascript
const options = SharePointBatch.GetSharePointOptions();
const batch = new SharePointBatch(options);
batch.add(new SharePointBatch.Changeset({ method: 'POST', url: '/_api/ContextInfo' }));
batch.add(new SharePointBatch.Changeset({ method: 'GET', url: '/_api/Site', params: { '$select': 'Id, Url, ReadOnly, WriteLocked' } }));
batch.add(new SharePointBatch.Changeset({ method: 'GET', url: '/_api/Web', params: { '$select': 'Id, Title, WebTemplate, Created' } }));
const response = await batch.send();
console.log(response.ok ? 'Done!' : 'Fail!', response.results);
```

### Documentation
Full documentation can be created by building the project and looking in the `docs` folder.

## Scripts
- `npm run build`
- `npm run build-src`
- `npm run build-docs`

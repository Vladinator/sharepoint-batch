# SharePointBatch
Inspiration drawn from the MSDN article ["Make batch requests with the REST APIs"](https://msdn.microsoft.com/en-us/library/office/dn903506.aspx)

This Javascript library is supposed to help developers efficiently build batch jobs and parsing the response from the server. The reason I made this was that the existing code provided by Microsoft seemed a bit underwhelming and the examples were too specific. There is a need for a general purpose library to utilize the API in existing projects.

Please note that this library is a work in progress. It's a mix between a research project and experimenting writing typescript to compile into JavaScript to then be used in other projects.

## Compatibility
Built to target ES5 compatible browsers.

## API
Full documentation can be created by building the project and looking in the `docs` folder.

```javascript
// the options object contains the properties `url` (hostweb) and `digest` (security)
const options = SharePointBatch.GetSharePointOptions();
// create a batch instance
const batch = new SharePointBatch(options);
// append jobs to the batch
batch.addChangeset(new SharePointBatch.Changeset({ method: 'POST', url: '/_api/ContextInfo' }));
batch.addChangeset(new SharePointBatch.Changeset({ method: 'GET', url: '/_api/Site', params: { '$select': 'Id, Url, ReadOnly, WriteLocked' } }));
batch.addChangeset(new SharePointBatch.Changeset({ method: 'GET', url: '/_api/Web', params: { '$select': 'Id, Title, WebTemplate, Created' } }));
// query the server for results
batch.send({
    done: function(options, response, results) { console.warn('Done!', results); },
    fail: function(options, response, error) { console.error('Fail!', error); },
});
```

## Scripts
- `npm run build`
- `npm run build-src`
- `npm run build-docs`

import SharePointBatch from './sharepoint';

((window: any) => {

    window.SharePointBatch = SharePointBatch;

    // DEBUG:
    (async () => {
        const options = SharePointBatch.GetSharePointOptions();
        if (!options)
            return console.error('This code can only run on a SharePoint site.');
        const spb = new SharePointBatch(options);
        window.SharePointBatchDebug = spb;
        /* DEBUG:
        SharePointBatchDebug._options.digest = '';
        SharePointBatchDebug.addChangeset(new SharePointBatch.Changeset({ url: '/_api/Site', finally: (options, changeset) => console.warn(changeset.getResponsePayload()) }));
        SharePointBatchDebug.addChangeset(new SharePointBatch.Changeset({ method: 'POST', url: '/_api/ContextInfo', finally: (options, changeset) => console.warn(changeset.getResponsePayload()) }));
        SharePointBatchDebug.addChangeset(new SharePointBatch.Changeset({ url: '/_api/Web', finally: (options, changeset) => console.warn(changeset.getResponsePayload()) }));
        response = await SharePointBatchDebug.send({ done: (options, response, payload) => console.warn('done', payload), fail: (options, response, payload, status, statusText) => console.warn('fail', payload, status, statusText) });
        // */
    })();

})(window);

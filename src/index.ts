import SharePointBatch from './sharepoint';

(async () => {

    const options = SharePointBatch.GetSharePointOptions();

    if (!options)
        return console.error('This code can only run on a SharePoint site.');

    const spb = new SharePointBatch(options);

    // DEBUG:
    (window as any).spb = spb;
    console.warn(await spb.GetWeb({
        '$expand': 'Lists,Webs',
        '$select': 'Id,Title,Lists/Id,Lists/Title,Webs/Id,Webs/Title',
    }));

})();

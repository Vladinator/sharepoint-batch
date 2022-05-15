import { Changeset, SharePointBatch } from './sharepoint';

((window: any, SharePointBatch: any) => {
    SharePointBatch.Changeset = Changeset;
    window.SharePointBatch = SharePointBatch;
})(window, SharePointBatch);

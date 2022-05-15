import { Changeset, BatchJob, SharePointBatch } from './sharepoint';

((window: any, SharePointBatch: any) => {
    SharePointBatch.Changeset = Changeset;
    SharePointBatch.BatchJob = BatchJob;
    window.SharePointBatch = SharePointBatch;
})(window, SharePointBatch);

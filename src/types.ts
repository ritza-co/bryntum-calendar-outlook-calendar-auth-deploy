import { Model, ProjectConsumer, ProjectModelMixin, Store } from '@bryntum/calendar';

export interface SyncDataParams {
  source: typeof ProjectConsumer | any;
  project: typeof ProjectModelMixin | any;
  store: Store;
  action: 'remove' | 'removeAll' | 'add' | 'clearchanges' | 'filter' | 'update' | 'dataset' | 'replace';
  record: Model;
  records: Model[];
  changes: object
}
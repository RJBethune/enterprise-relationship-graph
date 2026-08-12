import { SPHttpClient } from '@microsoft/sp-http';
import { SpRest, ISpTransport } from './sp/SpRest';
import { SpHttpTransport } from './sp/SpHttpTransport';
import { SpProvisioningService, IProvisioningService } from './sp/SpProvisioningService';
import { DocumentGraphStore } from './sp/DocumentGraphStore';
import { ItemGraphStore } from './sp/ItemGraphStore';
import { PresenceService } from './sp/PresenceService';
import { IGraphStore, StorageMode } from './IGraphStore';

/**
 * Composition root.
 *
 * The store is chosen per project, from the project's own `ErgStorageMode` column, so a
 * site can run a small graph as a single document and a heavily co-edited one per item
 * without any configuration. Everything above this factory is written against
 * IGraphStore and never learns which it got.
 */
export interface IErgServices {
  sp: SpRest;
  provisioning: IProvisioningService;
  presence: PresenceService;
  storeFor(mode: StorageMode): IGraphStore;
}

export const createServices = (
  transport: ISpTransport,
  editorName: string,
  editorLogin: string = editorName
): IErgServices => {
  const sp = new SpRest(transport);
  const documentStore = new DocumentGraphStore(sp, editorName);
  const itemStore = new ItemGraphStore(sp, editorName);
  return {
    sp,
    provisioning: new SpProvisioningService(sp),
    presence: new PresenceService(sp, editorLogin, editorName),
    storeFor: (mode: StorageMode): IGraphStore => (mode === 'Items' ? itemStore : documentStore)
  };
};

export const createSharePointServices = (
  client: SPHttpClient,
  webUrl: string,
  editorName: string,
  editorLogin: string
): IErgServices => createServices(new SpHttpTransport(client, webUrl), editorName, editorLogin);

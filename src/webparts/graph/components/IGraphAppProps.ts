import { IErgServices } from '../../../services/ServiceFactory';

export interface IGraphAppProps {
  services: IErgServices;
  /** Display name of the signed-in user; stamped on saves. */
  editorName: string;
  /** True when the user can write to the project list — drives read-only mode. */
  canEdit: boolean;
  /** Project to open on load (web part property or ?project= on the URL). */
  initialProjectId: number | null;
  /** Seconds between remote-change polls. 0 disables polling. */
  pollSeconds: number;
  /** Milliseconds of quiet before an edit is written. */
  autosaveMs: number;
  /** Base URL for font assets; defaults to the project's CDN folder. */
  assetBaseUrl: string;
  /** Remembers the last opened project across sessions. */
  onProjectChanged(projectId: number): void;
}

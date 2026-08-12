import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { SPPermission } from '@microsoft/sp-page-context';
import {
  IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'GraphWebPartStrings';
import GraphApp from './components/GraphApp';
import { IGraphAppProps } from './components/IGraphAppProps';
import { createSharePointServices, IErgServices } from '../../services/ServiceFactory';

export interface IGraphWebPartProps {
  projectId: number;
  viewHeight: number;
  pollSeconds: number;
  autosaveMs: number;
  assetBaseUrl: string;
}

const LAST_PROJECT_KEY = 'erg.lastProject';

export default class GraphWebPart extends BaseClientSideWebPart<IGraphWebPartProps> {
  private services: IErgServices | undefined;
  private editorName: string = '';
  private canEdit: boolean = false;

  protected onInit(): Promise<void> {
    const page = this.context.pageContext;
    this.editorName = page.user.displayName || page.user.email || 'Unknown';
    // Writing a graph means adding and editing list items; running setup additionally
    // needs manageLists, which the setup panel checks for separately.
    this.canEdit = page.web.permissions.hasPermission(SPPermission.addListItems) &&
      page.web.permissions.hasPermission(SPPermission.editListItems);

    this.services = createSharePointServices(
      this.context.spHttpClient,
      page.web.absoluteUrl,
      this.editorName,
      // loginName is the stable identity; display names collide and change.
      page.user.loginName || page.user.email || this.editorName,
      page.user.email || ''
    );
    return super.onInit();
  }

  /**
   * `?project=<id>` wins over the saved and configured defaults, so a link to a
   * specific graph always opens that graph — the property is the fallback, not the rule.
   */
  private resolveInitialProject(): number | null {
    try {
      const fromUrl = new URLSearchParams(window.location.search).get('project');
      if (fromUrl && !isNaN(Number(fromUrl))) { return Number(fromUrl); }
    } catch { /* older browsers */ }
    if (this.properties.projectId) { return this.properties.projectId; }
    try {
      const remembered = window.localStorage.getItem(LAST_PROJECT_KEY);
      if (remembered && !isNaN(Number(remembered))) { return Number(remembered); }
    } catch { /* private mode */ }
    return null;
  }

  public render(): void {
    if (!this.services) { return; }

    const element: React.ReactElement<IGraphAppProps> = React.createElement(GraphApp, {
      services: this.services,
      editorName: this.editorName,
      canEdit: this.canEdit,
      initialProjectId: this.resolveInitialProject(),
      pollSeconds: this.properties.pollSeconds === undefined ? 8 : this.properties.pollSeconds,
      autosaveMs: this.properties.autosaveMs === undefined ? 2000 : this.properties.autosaveMs,
      viewHeight: this.properties.viewHeight === undefined ? 0 : this.properties.viewHeight,
      // Blank is the DEV default (public font CDNs). Production sites set this to the
      // project's CDN folder — see loadEngineFonts.
      assetBaseUrl: this.properties.assetBaseUrl || '',
      onProjectChanged: (projectId: number): void => {
        try { window.localStorage.setItem(LAST_PROJECT_KEY, String(projectId)); } catch { /* private mode */ }
      }
    });

    // The graph is a full-canvas application, not a card: give it the height it needs
    // even when the page section does not offer one.
    this.domElement.style.height = '100%';
    this.domElement.style.minHeight = `${this.properties.viewHeight > 0 ? this.properties.viewHeight : 640}px`;

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    // Unmounting runs GraphApp.componentWillUnmount, which stops the engine's render
    // loop and removes its document listeners. Without that, warm SPA navigation away
    // from the page would leave a canvas loop running for the rest of the session.
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataGroupName,
              groupFields: [
                PropertyPaneSlider('projectId', {
                  label: strings.ProjectIdFieldLabel, min: 0, max: 200, step: 1
                }),
                PropertyPaneLabel('projectIdHelp', {
                  text: 'Leave at 0 to reopen whichever graph this browser used last. A ?project=<id> link always wins.'
                })
              ]
            },
            {
              groupName: strings.AppearanceGroupName,
              groupFields: [
                PropertyPaneSlider('viewHeight', {
                  label: strings.ViewHeightFieldLabel, min: 0, max: 2000, step: 20
                }),
                PropertyPaneLabel('viewHeightHelp', {
                  text: '0 fits the graph to the space left in the window. Set a height to force ' +
                    'a taller canvas — useful on small screens, where filling the window leaves ' +
                    'very little room. The page scrolls to reach it.'
                })
              ]
            },
            {
              groupName: strings.SyncGroupName,
              groupFields: [
                PropertyPaneSlider('pollSeconds', {
                  label: strings.PollSecondsFieldLabel, min: 0, max: 60, step: 1
                }),
                PropertyPaneSlider('autosaveMs', {
                  label: strings.AutosaveFieldLabel, min: 500, max: 10000, step: 250
                }),
                PropertyPaneLabel('syncHelp', {
                  text: 'Polling only runs while the tab is visible. Saving always flushes when the tab is hidden or closed.'
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneTextField('assetBaseUrl', { label: strings.AssetBaseUrlFieldLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloWorldApplicationCustomizerStrings';

// Tried presets instead of selective imports per early suggestion
import { sp } from "@pnp/sp/presets/all";
import { IHubSiteInfo } from "@pnp/sp/hubsites";
import { IItem } from "@pnp/sp/items";


const LOG_SOURCE: string = 'HelloWorldApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHelloWorldApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HelloWorldApplicationCustomizer
  extends BaseApplicationCustomizer<IHelloWorldApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    console.log(`Hello from ${strings.Title}:\n\n${message} 4`);

    await this.getHubSiteFromService(this.context?.pageContext?.legacyPageContext?.hubSiteId);
    await this.getPageListItem((this.context?.pageContext?.list.id as any)._guid, this.context?.pageContext?.listItem.id);

    return Promise.resolve();
  }

  private getHubSiteFromService = async (hubSiteId: string) => {
    try {
      console.log('Getting HUB SITE INFO');
      const hubSiteInfo: IHubSiteInfo = await sp.hubSites.getById(hubSiteId)();
      console.log('IHubSiteInfo from service - ', hubSiteInfo);
      return hubSiteInfo;
    }
    catch (error) {
      console.error('PnP call error - exception getting hub site - ', error);
      return undefined;
    }
  }

  private getPageListItem = async (listId: string, listItemId: number): Promise<IItem> => {
    try {
      let pageListItem: IItem = await sp.web.lists.getById(listId).items.getById(listItemId).usingCaching().get();
      console.info('PageListItem - ', pageListItem);

      return pageListItem;
    }
    catch (error) {
      console.warn('PnP call error list item - ', error);
    }
  }
}

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'HelloWorldApplicationCustomizerStrings';
import { sp } from "@pnp/sp/presets/all";
import { IHubSiteInfo } from "@pnp/sp/hubsites";

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

    console.log(`Hello from ${strings.Title}:\n\n${message}`);

    await this.getHubSiteFromService(this.context?.pageContext?.legacyPageContext?.hubSiteId);

    return Promise.resolve();
  }

  private getHubSiteFromService = async (hubSiteId: string) => {
    try {
      const hubSiteInfo: IHubSiteInfo = await sp.hubSites.getById(hubSiteId)();
      console.info('IHubSiteInfo from service - ', hubSiteInfo);
      return hubSiteInfo;
    }
    catch (error) {
      console.error('PnP call error - exception getting hub site - ', error);
      return undefined;
    }
  }
}

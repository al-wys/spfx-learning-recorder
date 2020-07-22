import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { SPHttpClient } from "@microsoft/sp-http";

import * as strings from 'LearningRecorderApplicationCustomizerStrings';

const LOG_SOURCE: string = 'LearningRecorderApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ILearningRecorderApplicationCustomizerProperties {
  // This is an example; replace with your own property
  recordListTitle: string;
  verificationPropertyName: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class LearningRecorderApplicationCustomizer
  extends BaseApplicationCustomizer<ILearningRecorderApplicationCustomizerProperties> {

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const pageInfo = await this.verifyPage();
    if (pageInfo.isLearningPage) {
      Log.info(LOG_SOURCE, "This is a learning page");
      console.log("This is a learning page");

      await this.addLearningRecord(pageInfo.title);
      Log.info(LOG_SOURCE, "Learning record is added");
      console.log("Learning record is added");
    } else {
      Log.info(LOG_SOURCE, "This is not a learning page");
      console.log("This is not a learning page");
    }

    // return Promise.resolve();
  }

  private async verifyPage(): Promise<{ title?: string, isLearningPage: boolean }> {
    const propertiesResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${location.pathname}')/ListItemAllFields?$select=Title,${this.properties.verificationPropertyName}`,
      SPHttpClient.configurations.v1);
    const properties: any = await propertiesResponse.json();

    if (properties[this.properties.verificationPropertyName]) {
      return { title: properties.Title, isLearningPage: true };
    } else {
      return { isLearningPage: false };
    }
  }

  private async getLearningRecordListInfo(): Promise<{ apiUrl: string, itemEntityTypeFullName: string, userId: number }> {
    const hubSiteInfoResponse = await this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/HubSiteData`, SPHttpClient.configurations.v1);
    const hubSiteInfo: { value: string } = await hubSiteInfoResponse.json();
    const hubSite: { url: string } = JSON.parse(hubSiteInfo.value);

    const infoArray = await Promise.all([
      this.context.spHttpClient.get(`${hubSite.url}/_api/lists/GetByTitle('${encodeURIComponent(this.properties.recordListTitle)}')`, SPHttpClient.configurations.v1),
      this.context.spHttpClient.get(`${hubSite.url}/_api/web/currentUser?$select=Id`, SPHttpClient.configurations.v1)
    ]);

    const recordListInfoResponse = infoArray[0];
    const recordListInfo: { Id: string, ListItemEntityTypeFullName: string } = await recordListInfoResponse.json();

    const userInfoResponse = infoArray[1];
    const userInfo: { Id: number } = await userInfoResponse.json();

    return { apiUrl: `${hubSite.url}/_api/web/lists(guid'${recordListInfo.Id}')`, itemEntityTypeFullName: recordListInfo.ListItemEntityTypeFullName, userId: userInfo.Id };
  }

  private async addLearningRecord(title: string): Promise<void> {
    const listInfo = await this.getLearningRecordListInfo();
    const data = {
      "Title": title,
      "UserId": listInfo.userId,
      "URL": {
        "Description": title,
        "Url": location.origin + location.pathname
      }
    };

    await this.context.spHttpClient.post(listInfo.apiUrl + "/items", SPHttpClient.configurations.v1, {
      body: JSON.stringify(data),
      headers: [
        ["accept", "application/json"]
      ]
    });
  }
}

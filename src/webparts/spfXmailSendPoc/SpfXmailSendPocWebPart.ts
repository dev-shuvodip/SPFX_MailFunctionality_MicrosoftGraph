import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";


import styles from './SpfXmailSendPocWebPart.module.scss';
import * as strings from 'SpfXmailSendPocWebPartStrings';

export interface ISpfXmailSendPocWebPartProps {
  description: string;
}

export default class SpfXmailSendPocWebPart extends BaseClientSideWebPart<ISpfXmailSendPocWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <div class="${styles.spfXmailSendPoc}">
      <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="ms-Grid-col ms-u-sm12 block">
             <label class="ms-Label ms-Grid-col ms-u-sm4 block">To</label>           
             <input id="emailID" class="ms-TextField-field  ms-Grid-col ms-u-sm8 block" type="text" value="" placeholder="">
             </div>
          </div>

          <div class="${styles.row}">
            <div class="ms-Grid-col ms-u-sm12 block">
            <label class="ms-Label ms-Grid-col ms-u-sm4 block">Subject</label>
            <input id="oSubjt" class="ms-TextField-field ms-Grid-col ms-u-sm8 block" type="text" value="" placeholder="">
            </div>
          </div>

          <div class="${styles.row}">
            <div class="ms-Grid-col ms-u-sm12 block">
            <label class="ms-Label ms-Grid-col ms-u-sm4 block">Email Content</label>
            <textarea id="emailContent" class="ms-TextField-field  ms-Grid-col ms-u-sm8 block"></textarea>
            </div>
            </div>
          <div class="${styles.row}">
            <button class="${styles.button} sendEmail-Button ms-Grid-col ms-u-sm12 block" id="send_mail">
              <span class="${styles.label}">Send Email</span>
            </button>
          </div>        
      </div>
    </div>`;

    this._bindEvents();
  }

  private _bindEvents() {
    document.getElementById('send_mail').addEventListener("click", (e: Event) => this._sendEmail());
  }

  private _sendEmail(): void {
    console.log("-- inside _sendEmail() --");
    var emailID: string = (<HTMLInputElement>document.getElementById("emailID")).value;
    var oSubject: string = (<HTMLInputElement>document.getElementById("oSubjt")).value;
    var oBody: string = (<HTMLInputElement>document.getElementById("emailContent")).value;

    const body: any = {
      message: {
        subject: oSubject,
        body: {
          contentType: "HTML",
          content: oBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: emailID
            }
          }
        ]
      },
      saveToSentItems: false
    };


    this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client
        .api('/me/sendMail')
        .post(body, (err, res) => {
          console.log("-- post res : " + JSON.stringify(res));
          if (err) {
            alert("Error")
          } else {
            alert("Email Sent");
          }
        });

    });
  }


  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import {
  MSGraphClient,
  AadTokenProvider,
  AadHttpClient,
  HttpClientResponse,
  HttpClient
} from '@microsoft/sp-http';

//import MSGraph typings (facultatif mais utile)
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import styles from './MsGraphClientDemoWebPart.module.scss';
import * as strings from 'MsGraphClientDemoWebPartStrings';

export interface IMsGraphClientDemoWebPartProps {
  description: string;
}

// Définition de la réponse du MSGraph
export interface IDriveItems {
  '@odata.context': string;
  value: MicrosoftGraph.DriveItem[]; // Collection de DriveItem
}

export default class MsGraphClientDemoWebPart extends BaseClientSideWebPart<IMsGraphClientDemoWebPartProps> {

  public render(): void {
    this._asyncRender();
  }

  // Instructions
  // * Décommenter chaque partie que vous souhaitez tester (cf. "décommenter la ligne ci-dessous pour tester")
  // * Utilisez le Workbench de votre site SharePoint Online pour tester
  // * Pour tester la partie Permissions, il faut packager la solution et la déployer
  private async _asyncRender() {
    let recentDocuments, recentDocumentItems;

    // --- Appel MSGraph avec aadHttpClient + async/await
    // --- Affichage : console (F12)
    // To do : décommenter la ligne ci-dessous pour tester
    // this._queryWithAadHttpClient();

    // --- Récupération d'un access token et appel MSGraph avec le token reçu
    // --- Affichage : console (F12)
    // To do : décommenter la ligne ci-dessous pour tester
    // this._getToken();

    // --- Appel MSGraph avec msGraphClient + promises
    // --- Affichage : dans la web part
    try {
      recentDocumentItems = await this._queryWithMSGraphClient();
    }
    catch {
      recentDocumentItems = null;
    }

    // Liste des documents récents à afficher dans la web part (rendu HTML)
    if (recentDocumentItems && recentDocumentItems.length > 0) {
      recentDocumentItems = recentDocumentItems.map((rd) => {
        return "<li><a href=" + rd.webUrl + " target='_blank'>" + rd.name + "</li></a>";
      });

      recentDocuments = "<ul>";
      let recentDocumentItemsLength = recentDocumentItems.length;
      for (let i = 0; i < recentDocumentItemsLength; i++) {
        recentDocuments += recentDocumentItems[i];
      }
      recentDocuments += "</ul>";
    }
    else if (recentDocumentItems && recentDocumentItems.length === 0) {
      recentDocuments = "Pas de documents";
    }
    else {
      recentDocuments = "Ouvrir la console de débogage (F12) pour voir les logs.";
    }

    // Rendu final
    this.domElement.innerHTML = `
      <div class="${ styles.msGraphClientDemo}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Microsoft Graph Demo</span>
              <h3>Liste des documents récents</h3>
                ${recentDocuments}
            </div>
          </div>
        </div>
      </div>`;
  }

  // Récupération d'un access token et interrogation du MSGraph via fetch
  private async _getToken() {
    const tokenProvider: AadTokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();

    // Récupération d'un access token pour pouvoir interroger le Microsoft Graph
    const token: string = await tokenProvider.getToken("https://graph.microsoft.com");
    console.log("token", token);

    // Url de l'endpoint pour récupérer les fichiers récents de l'utilisateur connecté
    const endpointUrl = "https://graph.microsoft.com/v1.0/me/drive/recent";

    // Interrogation du Microsoft Graph sans aadHttpClient ou msGraphClient
    const response: Response = await fetch(endpointUrl, {
      method: 'GET',
      headers: {
        "Authorization": "Bearer " + token,
        "Accept": "application/json"
      }
    });

    if (!response.ok) {
      console.log("[MsGraphClientDemoWebPart._getToken] Erreur : la requête au Microsoft Graph n'a pas pu être executée correctement");
      return;
    }

    const driveItemsJson: IDriveItems = await response.json();
    const driveItems: MicrosoftGraph.DriveItem[] = driveItemsJson.value;
    console.log("[MsGraphClientDemoWebPart._getToken] driveItems : ", driveItems);
  }

  // Appel Microsoft Graph avec aadHttpClient et async/await
  private async _queryWithAadHttpClient() {
    const client: AadHttpClient = await this.context.aadHttpClientFactory.getClient('https://graph.microsoft.com');
    const response: HttpClientResponse = await client.get('https://graph.microsoft.com/v1.0/me/drive/recent', AadHttpClient.configurations.v1);

    if (!response.ok) {
      console.log("[MsGraphClientDemoWebPart._queryWithAadHttpClient] Erreur : la requête au Microsoft Graph n'a pas pu être executée correctement");
      return;
    }

    const jsonResponse: IDriveItems = await response.json();
    console.log("[MsGraphClientDemoWebPart._queryWithAadHttpClient] jsonResponse :", jsonResponse.value);
  }

  // Appel Microsoft Graph avec msGraphClient et promises
  private async _queryWithMSGraphClient() {
    return new Promise((resolve, reject) => {
      let _recentDocuments: Object[] = [];
      this.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient): void => {
          client.api('/me/drive/recent')
            .select('name, weburl')
            .get((err, res: IDriveItems, rawResponse?: any) => {
              if (err) {
                // Gestion des erreurs
                if (err.message) {
                  console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Erreur lors de l'execution de la requête : ", err.message);
                }
                else { // Erreur générique
                  console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Erreur lors de l'execution de la requête.");
                }

                reject(err);
                return;
              }

              // On vérifie s'il y a au moins un résultat
              if (res && res.value && res.value.length > 0) {
                const files: MicrosoftGraph.DriveItem[] = res.value;
                const filesCount = files.length;

                console.log(`Cas où il y a des fichiers récents : ${filesCount} éléments`);
                console.log('[MsGraphClientDemoWebPart._queryWithMSGraphClient] res.value :', files);

                // Listing des fichiers récents sur OneDrive
                for (let i = 0; i < filesCount; i++) {
                  const f: MicrosoftGraph.DriveItem = files[i];
                  console.log(f.name);
                  _recentDocuments.push(f);
                }
                resolve(_recentDocuments);
              }
              else {
                console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Il n'y a pas de fichiers récents");
                resolve(_recentDocuments);
              }
            });
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
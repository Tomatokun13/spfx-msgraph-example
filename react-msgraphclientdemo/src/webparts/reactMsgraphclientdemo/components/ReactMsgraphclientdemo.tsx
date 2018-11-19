import * as React from 'react';
import styles from './ReactMsgraphclientdemo.module.scss';
import { IReactMsgraphclientdemoProps } from './IReactMsgraphclientdemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { FieldUrlRenderer } from "@pnp/spfx-controls-react/lib/FieldUrlRenderer";

export interface IReactMsgraphclientdemoState {
  recentDocuments: MicrosoftGraph.DriveItem[];
  loading: boolean;
}

// Définition de la réponse du MSGraph
export interface IDriveItems {
  '@odata.context': string;
  value: MicrosoftGraph.DriveItem[]; // Collection de DriveItem
}

const _columns: IColumn[] = [
  {
    key: 'column1',
    name: 'Name',
    fieldName: 'name',
    minWidth: 200,
    maxWidth: 300,
    isResizable: true,
    onRender: (item: MicrosoftGraph.DriveItem) => {
      return <FieldUrlRenderer text={item.name} url={item.webUrl} />;
    }
  },
  {
    key: 'column2',
    name: 'Url',
    fieldName: 'webUrl',
    minWidth: 200,
    maxWidth: 500,
    isResizable: true
  }
];



export default class ReactMsgraphclientdemo extends React.Component<IReactMsgraphclientdemoProps, IReactMsgraphclientdemoState> {
  public state = {
    recentDocuments: [],
    loading: true
  };

  public render(): React.ReactElement<IReactMsgraphclientdemoProps> {
    const { recentDocuments, loading } = this.state;

    return (
      <div className="styles.msGraphClientDemo">
        <div className="styles.container">
          <div className="styles.row">
            <div className="styles.column">
              <span className="styles.title">Microsoft Graph Demo</span>
              <h3>Liste des documents récents</h3>
              {
                loading ?
                  <Spinner label="Chargement des données ..." />
                :
                  <DetailsList
                  items={recentDocuments}
                  columns={_columns}
                  setKey="set"
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  selectionPreservedOnEmptyClick={true}
                  ariaLabelForSelectionColumn="Toggle selection"
                  ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                />
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  public componentDidMount(): void {
    this.props.graphClient
      .api("me/drive/recent")
      .select('name, weburl') // on sélectionne les métadonnées name et webUrl
      .get((err, res: IDriveItems, rawResponse?: any) => {
        if (err) {
          // Gestion des erreurs
          if (err.message) {
            console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Erreur lors de l'execution de la requête : ", err.message);
          }
          else { // Erreur générique
            console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Erreur lors de l'execution de la requête.");
          }
          this.setState({
            recentDocuments: [],
            loading: false
          });
          return;
        }

        // On vérifie s'il y a au moins un résultat
        if (res && res.value && res.value.length > 0) {
          const files: MicrosoftGraph.DriveItem[] = res.value;
          const filesCount = files.length;

          console.log(`Cas où il y a des fichiers récents : ${filesCount} éléments`);
          console.log('[MsGraphClientDemoWebPart._queryWithMSGraphClient] res.value :', files);
          this.setState({
            recentDocuments: files,
            loading: false
          });
        }
        else {
          console.log("[MsGraphClientDemoWebPart._queryWithMSGraphClient] Il n'y a pas de fichiers récents");
          this.setState({
            recentDocuments: [],
            loading: false
          });
        }
      });
  }
}

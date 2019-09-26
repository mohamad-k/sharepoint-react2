import * as React from 'react';
import { sp } from '@pnp/sp';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, CheckboxVisibility } from 'office-ui-fabric-react/lib/DetailsList';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import ContentTypeForm from './ContentTypeForm';
import IDocumentModel from '../models/IDocumentModel';

export interface ITaskViewProps {
  // documentId: number;
}

export interface ITaskViewState {
  documents: IDocumentModel[];
  filteredDocuments: IDocumentModel[];
  isModalOpen: boolean;
}

export default class TaskView extends React.Component<ITaskViewProps, ITaskViewState> {
  public state: ITaskViewState = {
    documents: new Array<IDocumentModel>(),
    filteredDocuments: new Array<IDocumentModel>(),
    isModalOpen: false
  };

  constructor() {
    super();

    this.onChanged = this.onChanged.bind(this);
    this.toggleDialog = this.toggleDialog.bind(this);
    this.onRender = this.onRender.bind(this);
  }

  public componentDidMount(): void {
    sp.web.lists
      .getByTitle('DxpDokument')
      .items.get()
      .then((documents: IDocumentModel[]) => {
        this.setState({ documents, filteredDocuments: documents });
        console.log(documents);
      });
  }

  public render(): JSX.Element {
    return (
      <div>
        <h1>Unveröffentlichte Dokumente</h1>
        <CommandBar
          isSearchBoxVisible={false}
          items={[{ key: 'filter', onRender: this.onRender }]}
          farItems={[{ key: 'new', name: 'Neues Dokument', iconProps: { iconName: 'Add' }, onClick: this.toggleDialog }]}
        />

        <DetailsList
          items={this.state.filteredDocuments}
          columns={[{ key: 'name', name: 'Dokumentenname', fieldName: 'Title', minWidth: 100 }]}
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick={true}
          checkboxVisibility={CheckboxVisibility.hidden}
        />

        <Dialog title="Neues Dokument" isOpen={this.state.isModalOpen} onDismiss={this.toggleDialog}>
          <ContentTypeForm contentTypeId="0x01000AC26F56A3A87847AA3E9637D33EBCEA" />
          <DialogFooter>
            <PrimaryButton onClick={this.toggleDialog} text="Save" />
            <DefaultButton onClick={this.toggleDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  private onChanged(text: string): void {
    if (text.length === 0) {
      console.log('Empty');
      this.setState({ filteredDocuments: this.state.documents });
    } else {
      console.log(text);
      const filteredDocuments: IDocumentModel[] = [];

      for (const document of this.state.documents) {
        if (document.Title.toLowerCase().indexOf(text.toLowerCase()) > -1) {
          filteredDocuments.push(document);
          console.log('pushed');
        }
      }
      this.setState({ filteredDocuments });
    }
  }

  private toggleDialog(e): void {
    e.preventDefault();
    this.setState({ isModalOpen: !this.state.isModalOpen });
  }

  private onRender(): JSX.Element {
    return <TextField placeholder="Filter" onChanged={this.onChanged} />;
  }
}


// enum 
export enum Views {
    BOOK_VIEW = 'bookOverview',
    TASK_VIEW = 'taskView',
    DOCUMENT_VIEW = 'documentView',
    ADM_BOOK_VIEW = 'admBookView',
    ADM_Document_VIEW = 'admDocumentsView',
    Dashboard_VIEW = 'dashboardView'
  }
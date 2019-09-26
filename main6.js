import * as React from 'react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { sp } from '@pnp/sp';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

import { Errors } from '../ErrorEnum';
import IDocumentModel from '../models/IDocumentModel';

export interface IDocumentViewProps {
  documentId: number;
}

export interface IDocumentViewState {
  document: IDocumentModel;
  error: string;
}

export default class DocumentView extends React.Component<IDocumentViewProps, IDocumentViewState> {
  public state: IDocumentViewState = {
    document: undefined,
    error: undefined
  };
  constructor() {
    super();

    this.fetchDocument = this.fetchDocument.bind(this);
    this.errorMessage = this.errorMessage.bind(this);
  }
  public componentWillReceiveProps(nextProps: IDocumentViewProps): void {
    this.fetchDocument(nextProps.documentId);
  }
  public componentDidMount(): void {
    this.fetchDocument(this.props.documentId);
  }

  public render(): JSX.Element {
    if (this.state.document) {
      return (
        <div>
          <h1>
            {this.state.document.Title} (#{this.state.document.ID})
          </h1>

          <div dangerouslySetInnerHTML={{ __html: this.state.document.DxpContent }} />
        </div>
      );
    } else if (this.state.error) {
      return this.errorMessage();
    } else {
      return <Spinner label="Bitte warten" />;
    }
  }
  private errorMessage(): JSX.Element {
    if (this.state.error) {
      console.log(this.state.error);
      return (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {Errors.NoDocumentFound}
        </MessageBar>
      );
    }
  }
  private fetchDocument(documentId: number): void {
    sp.web.lists
      .getByTitle('DxpDokument')
      .items.getById(documentId)
      .select('*', 'Author/ID', 'Author/Title', 'Author/Name')
      .expand('Author')
      .get()
      .then((document: IDocumentModel) => {
        this.setState({ document, error: undefined });
      })
      .catch(error => {
        this.setState({ error, document: undefined });
        console.log('no document');
      });
  }
}

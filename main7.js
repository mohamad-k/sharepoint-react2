import * as React from 'react';
import { sp } from '@pnp/sp';

import Navbar from './Navbar';
import TaskView from './TaskView';
import DocumentView from './DocumentView';
import SearchView from './SearchView';
import IDocumentModel from '../models/IDocumentModel';
import { Errors } from '../ErrorEnum';
import BookOverview from './BookOverview';
import BookSettingsView from './admin/BookSettingsView';
import DashboardView from './DashboardView';
import { Views } from './ViewEnum';
import styles from './Docxpert.module.scss';

interface ISessionItem {
  view: string;
  param?: object;
}

export interface IDocxpertProps {
  description: string;
}

export interface IDocxpertState {
  view: string;
  documentId: number;
  error: string;
  documents: IDocumentModel[];
}

export default class Docxpert extends React.Component<IDocxpertProps, IDocxpertState> {
  public state: IDocxpertState = {
    view: 'bookOverview',
    documentId: undefined,
    documents: [],
    error: undefined
  };

  constructor() {
    super();

    this.changeView = this.changeView.bind(this);
    this.loadView = this.loadView.bind(this);
    this.searchFunc = this.searchFunc.bind(this);
  }

  public componentWillMount(): void {
    const itemJson: string = sessionStorage.getItem('docxpert');

    if (itemJson) {
      const item: { view: string; param: { documentId: number } } = JSON.parse(itemJson);
      this.setState({ view: item.view, documentId: item.param.documentId });
    }
  }

  public render(): JSX.Element {
    return (
      <div className={styles.docxpert}>
        <Navbar search={this.searchFunc} changeView={this.changeView} />
        {this.loadView()}
      </div>
    );
  }

  private changeView(view: string, documentId?: number): void {
    this.setState({ view, documentId });
  }

  private loadView(): React.ReactElement {
    const viewParam: ISessionItem = { view: this.state.view, param: { documentId: this.state.documentId } };
    sessionStorage.setItem('docxpert', JSON.stringify(viewParam));

    switch (this.state.view) {
      case Views.Dashboard_VIEW:
        return <DashboardView changeView={this.changeView} />;
      case Views.BOOK_VIEW:
        return <BookOverview changeView={this.changeView} />;
      case Views.TASK_VIEW:
        return <TaskView />;
      case Views.DOCUMENT_VIEW:
        return <DocumentView documentId={this.state.documentId} />;
      case 'searchView':
        return <SearchView changeView={this.changeView} error={this.state.error} documents={this.state.documents} />;
      case Views.ADM_BOOK_VIEW:
        return <BookSettingsView />;
      case Views.ADM_Document_VIEW:
      default:
        return <BookOverview changeView={this.changeView} />;
    }
  }
  private searchFunc(search: string): void {
    if (search.match(/#\d+/)) {
      const id: number = parseInt(search.substr(1), 10);
      this.changeView('documentView', id);
    } else {
      sp.web.lists
        .getByTitle('DxpDokument')
        .items.filter(`substringof('${search}',Title)`)
        .select('*', 'DxpBook/ID', 'DxpBook/Title')
        .expand('DxpBook')
        .get()
        .then(documents => {
          if (documents.length > 0) {
            this.setState({ documents });
            console.log(documents);
            this.changeView('searchView');
          } else {
            console.log('set error');
            this.setState({ documents: [], error: Errors.NoDocumentFound });
            this.changeView('searchView');
          }
        });
    }
  }
}

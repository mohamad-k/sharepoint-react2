import * as React from 'react';
import { sp } from '@pnp/sp';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { DetailsList, DetailsListLayoutMode, CheckboxVisibility, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import IBookModel from '../models/IBookModel';
import IDocumentModel from '../models/IDocumentModel';
import IDocumentListItem from '../models/IDocumentListItem';
import { Errors } from '../ErrorEnum';
import styles from './Docxpert.module.scss';

export interface IDashboardViewProps {
  changeView: Function;
}

export interface IDashboardViewState {
  error: string;
  books: IBookModel[];
  documents: IDocumentModel[];
  sortedBooks: boolean;
  sortedDocuments: boolean;
  documentsBookNameSorted: boolean;
  documentCreatedSorted: boolean;
  isDocumentNameSorted: boolean;
  isDocumentBookNameSorted: boolean;
  isDocumentCreatedSorted: boolean;
}

export default class DashboardView extends React.Component<IDashboardViewProps, IDashboardViewState> {
  public state: IDashboardViewState = {
    error: undefined,
    books: [],
    documents: [],
    sortedBooks: false,
    sortedDocuments: false,
    documentsBookNameSorted: false,
    documentCreatedSorted: false,
    isDocumentNameSorted: undefined,
    isDocumentBookNameSorted: undefined,
    isDocumentCreatedSorted: undefined
  };

  constructor() {
    super();

    this.fetchBooksAndDocuments = this.fetchBooksAndDocuments.bind(this);
    this.onRenderBookColumn = this.onRenderBookColumn.bind(this);
    this.onRenderDocumentColumn = this.onRenderDocumentColumn.bind(this);
    this.sortBooks = this.sortBooks.bind(this);
    this.sortDocuments = this.sortDocuments.bind(this);
    this.sortDocumentsBookName = this.sortDocumentsBookName.bind(this);
    this.sortDocumentsCreated = this.sortDocumentsCreated.bind(this);
  }

  public componentDidMount(): void {
    this.fetchBooksAndDocuments();
  }

  public render(): JSX.Element {
    return (
      <div className={styles.dashboardview}>
          <div className='ms-Grid'>
              <div className='ms-Grid-row'>
        <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6'>
            <p className='ms-font-xxl ms-fontWeight-semibold '>Neu veröffentlichte Dokumente</p>
          <DetailsList
            items={this.state.documents}
            onRenderItemColumn={this.onRenderDocumentColumn}
            checkboxVisibility={CheckboxVisibility.hidden}
            columns={[
              {
                name: 'Name',
                key: 'name',
                fieldName: 'Title',
                minWidth: 50,
                maxWidth: 100,
                isSorted: this.state.isDocumentNameSorted,
                isSortedDescending: this.state.sortedDocuments ? true : false,
                isResizable: true,
                onColumnClick: this.sortDocuments
              },
              {
                name: 'Buch',
                key: 'buch',
                fieldName: 'Title',
                minWidth: 50,
                maxWidth: 100,
                isSorted: this.state.isDocumentBookNameSorted,
                isSortedDescending: this.state.documentsBookNameSorted ? true : false,
                isResizable: true,
                onColumnClick: this.sortDocumentsBookName
              },
              { name: 'Version', key: 'version', fieldName: 'Version', minWidth: 50, maxWidth: 50, isResizable: true },
              { name: 'veröffentlicht am', key: 'created', fieldName: 'Created', minWidth: 50, maxWidth: 140, isResizable: true,
              isSorted: this.state.isDocumentCreatedSorted,
              isSortedDescending: this.state.documentCreatedSorted ? false : true,
              onColumnClick: this.sortDocumentsCreated },
              { name: 'Redaktionelle Änderung?', key: 'änderung', fieldName: 'änderung', minWidth: 50, isResizable: true }
            ]}
          />
        </div>
        <div className='ms-Grid-col ms-sm6 ms-md4 ms-lg6'>
        <p className='ms-font-xxl ms-fontWeight-semibold'>Meine Bücher</p>
          <DetailsList
            items={this.state.books}
            onRenderItemColumn={this.onRenderBookColumn}
            checkboxVisibility={CheckboxVisibility.hidden}
            columns={[
              // { name: 'ID', key: 'ID', fieldName: 'ID', minWidth: 10, maxWidth: 20 },
              {
                name: 'Name',
                key: 'name',
                fieldName: 'Title',
                minWidth: 50,
                isSorted: true,
                isSortedDescending: this.state.sortedBooks ? false : true,
                isResizable: true,
                onColumnClick: this.sortBooks
              },
              { name: 'Aktion', key: 'action', fieldName: undefined, minWidth: 100, isResizable: true }
            ]}
          />
        </div>
        </div>
        </div>
      </div>
    );
  }

  private fetchBooksAndDocuments(): void {
    sp.web.lists
      .getByTitle('DxpBuch')
      .items.orderBy('Title', true)
      .get()
      .then((books: IBookModel[]) => {
        if (books.length > 0) {
          this.setState({ books });
        }
        throw Errors.NoBookFound;
      });
    sp.web.lists
      .getByTitle('DxpDokument')
      .items.select('*', 'DxpBook/ID', 'DxpBook/Title')
      .expand('DxpBook')
      .orderBy('Created')
      .top(4)
      .get()
      .then((documents: IDocumentModel[]) => {
        documents.reverse();
        this.setState({ documents });
        console.log(documents)
      })
      .catch(error => {
        this.setState({ error });
      });
  }

  private onRenderDocumentColumn(item: any, index: number, column: IColumn): JSX.Element {
    switch (column.key) {
      case 'name':
        return (
          <Link className={styles.link} onClick={() => this.changeView(item.ID)}>
            {item.Title}
          </Link>
        );
      case 'buch':
        return <Link className={styles.link}>{item.DxpBook.Title}</Link>;
      case 'created':
        return <p>{item.Created}</p>;
      default:
        <Spinner label="Bitte warten" />;
        break;
    }
  }

  private onRenderBookColumn(book: any, index: number, column: IColumn): JSX.Element {
    switch (column.key) {
      case 'name':
        return <Link className={styles.link}>{book.Title}</Link>;
      default:
        <Spinner />;
        break;
    }
  }
  private changeView(itemId: number) {
    this.props.changeView('documentView', itemId);
  }
  private sortBooks() {
    this.state.books.reverse();
    this.setState({ sortedBooks: !this.state.sortedBooks });
  }
  private sortDocuments() {
    this.setState({ 
        sortedDocuments: !this.state.sortedDocuments, 
        isDocumentNameSorted: true, 
        isDocumentBookNameSorted: false , 
        isDocumentCreatedSorted: false});
    let sortingDocuments = this.state.documents.sort((a, b) => {
      if (!this.state.sortedDocuments) {
        if (b.Title.toUpperCase() < a.Title.toUpperCase()) {
          return -1;
        }if (b.Title.toUpperCase() > a.Title.toUpperCase()) {
            return 1;
          }
      }
    });
    this.setState({ documents: sortingDocuments });
    this.state.documents.reverse();
  }
  private sortDocumentsBookName() {
    this.setState({
         documentsBookNameSorted: !this.state.documentsBookNameSorted, 
         isDocumentBookNameSorted: true, 
         isDocumentNameSorted: false, 
         isDocumentCreatedSorted: false });
    let sortingDocuments = this.state.documents.sort((a, b) => {
      if (!this.state.documentsBookNameSorted) {
        if (b.DxpBook.Title.toUpperCase() < a.DxpBook.Title.toUpperCase()) {
          return -1;
        }
        if (b.DxpBook.Title.toUpperCase() > a.DxpBook.Title.toUpperCase()) {
          return 1;
        }
      }
    });
    this.setState({ documents: sortingDocuments });
    this.state.documents.reverse();
  }
  private sortDocumentsCreated(){
    this.setState({ 
        documentCreatedSorted: !this.state.documentCreatedSorted, 
        isDocumentBookNameSorted: false, 
        isDocumentNameSorted: false, 
        isDocumentCreatedSorted: true });
    let sortingDocuments = this.state.documents.sort((a, b) => {
      if (!this.state.documentCreatedSorted) {
        if (b.Created.toUpperCase() < a.Created.toUpperCase()) {
          return -1;
        }if (b.Created.toUpperCase() > a.Created.toUpperCase()) {
            return 1;
          }
      }
    });
    this.setState({ documents: sortingDocuments });
    this.state.documents.reverse();
  }
  private sorting(val1, val2){
    this.state.documents.sort((a, b) => {
        if (!this.state.documentCreatedSorted) {
          if (b.Created.toUpperCase() < a.Created.toUpperCase()) {
            return -1;
          }if (b.Title.toUpperCase() > a.Title.toUpperCase()) {
              return 1;
            }
        }
    })
  }
}

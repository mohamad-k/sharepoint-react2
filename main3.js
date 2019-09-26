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

export interface IBookOverviewProps {
  changeView: Function;
}

export interface IBookOverviewState {
  error: string;
  books: IBookModel[];
  selectedBook: number | undefined;
  documents: IDocumentListItem[];
}

export default class BookOverview extends React.Component<IBookOverviewProps, IBookOverviewState> {
  public state: IBookOverviewState = {
    error: undefined,
    books: [],
    selectedBook: undefined,
    documents: []
  };

  constructor() {
    super();

    this.errorMessage = this.errorMessage.bind(this);
    this.fetchBooksAndDocuments = this.fetchBooksAndDocuments.bind(this);
    this.renderItemColumn = this.renderItemColumn.bind(this);
    this.toggleChildren = this.toggleChildren.bind(this);
    this.changeState = this.changeState.bind(this);
  }

  public componentDidMount(): void {
    this.fetchBooksAndDocuments();
  }

  public render(): JSX.Element {
    return (
      <div className={styles.bookview}>
        <Dropdown
          label="Buch"
          selectedKey={this.state.selectedBook}
          onChanged={this.changeState}
          options={this.state.books.map(book => ({
            key: book.ID,
            text: book.Title
          }))}
        />

        <DetailsList
          items={this.state.documents}
          columns={[{ key: 'name', name: 'Dokumentenname', fieldName: 'Title', minWidth: 100 }]}
          layoutMode={DetailsListLayoutMode.justified}
          checkboxVisibility={CheckboxVisibility.hidden}
          onRenderItemColumn={this.renderItemColumn}
        />
        {this.errorMessage()}
        {this.isLoading()}
      </div>
    );
  }

  private errorMessage(): JSX.Element {
    if (this.state.error) {
      return (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {this.state.error}
        </MessageBar>
      );
    }
  }

  private isLoading(): JSX.Element {
    if ((!this.state.documents || this.state.documents.length === 0) && !this.state.error) {
      return <Spinner label="loading..." />;
    }
  }

  private fetchBooksAndDocuments(): void {
    sp.web.lists
      .getByTitle('DxpBuch')
      .items.orderBy('Order1', true)
      .get()
      .then((books: IBookModel[]) => {
        console.log(books);
        if (books.length > 0) {
          this.setState({ books, selectedBook: books[0].ID });
          return books[0].ID;
        }

        throw Errors.NoBookFound;
      })
      .then((bookId: number) => {
        sp.web.lists
          .getByTitle('DxpDokument')
          .items.filter(`DxpBookId eq ${bookId}`)
          .get()
          .then((documents: IDocumentModel[]) => {
            if (documents.length === 0) {
              this.setState({ error: Errors.NoDocumentFound });
              throw Errors.NoDocumentFound;
            }

            sp.web.lists
              .getByTitle('DxpDokument')
              .items.filter(this.childrenFilter(documents))
              .get()
              .then((children: IDocumentModel[]) => {
                const documentList: IDocumentListItem[] = [];

                for (const parent of documents) {
                  const item: IDocumentListItem = this.convertDocumentToItem(parent);

                  for (const child of children) {
                    if (child.DxpParentId === item.ID) {
                      item.children.push(child);
                    }
                  }

                  documentList.push(item);
                }

                this.setState({ documents: documentList });
              });
          });
      })
      .catch(error => {
        this.setState({ error });
      });
  }

  private convertDocumentToItem(document: IDocumentModel, node?: number) {
    const item = {} as IDocumentListItem;

    Object.keys(document).map(field => (item[field] = document[field]));

    item.children = [];
    item.isOpen = false;
    item.node = node ? node : 0;

    return item;
  }

  private childrenFilter(parents: IDocumentModel[]): string {
    let filter: string = '';

    for (let i: number = 0; i < parents.length; i++) {
      filter += `DxpParentId eq ${parents[i].ID}`;

      if (i < parents.length - 1) {
        filter += ' or ';
      }
    }

    return filter;
  }

  private renderItemColumn(item: IDocumentListItem, index: number, column: IColumn): JSX.Element {
    switch (column.key) {
      case 'name':
        if (item.children && item.children.length > 0) {
          return (
            <span style={{ 'padding-left': `${16 * item.node}px` }}>
              <IconButton
                style={{ height: 'inherit', paddingRight: '20px' }}
                iconProps={{ iconName: item.isOpen ? 'ChevronDown' : 'ChevronRight' }}
                onClick={() => {
                  this.toggleChildren(index);
                }}
              />
              <Link className={styles.link} onClick={() => this.changeView(item.ID)}>
                {item.Title}
              </Link>
            </span>
          );
        } else {
          return (
            <span style={{ 'padding-left': `${16 * item.node + 35}px` }}>
              <Link className={styles.link} onClick={() => this.changeView(item.ID)}>
                {item.Title}
              </Link>
            </span>
          );
        }
      default:
        return <Spinner label="I am definitely loading..." />;
    }
  }

  private toggleChildren(index: number): void {
    const documents: IDocumentListItem[] = this.state.documents.slice();

    if (documents[index].isOpen) {
      let deleteCount: number = 0;

      for (let i: number = index + 1; i < documents.length; i++) {
        if (documents[i].node <= documents[index].node) {
          break;
        }

        deleteCount++;
      }

      documents.splice(index + 1, deleteCount);
      this.setState({ documents });
    } else {
      sp.web.lists
        .getByTitle('DxpDokument')
        .items.filter(this.childrenFilter(documents[index].children))
        .get()
        .then((children: IDocumentModel[]) => {
          console.log(children);
          for (const parent of documents[index].children) {
            const item: IDocumentListItem = this.convertDocumentToItem(parent, documents[index].node);
            item.node++;

            for (const child of children) {
              if (child.DxpParentId === parent.ID) {
                item.children.push(child);
              }
            }

            documents.splice(index + 1, 0, item);
          }

          this.setState({ documents });
        });
    }

    documents[index].isOpen = !documents[index].isOpen;
  }
  private changeState(item: { key: number; text: string }): void {
    sp.web.lists
      .getByTitle('DxpDokument')
      .items.filter(`DxpBookId eq ${item.key} and DxpParentId eq null`)
      .get()
      .then((documents: IDocumentModel[]) => {
        if (documents.length === 0) {
          this.setState({ documents: [], selectedBook: item.key, error: Errors.NoDocumentFound });
          throw Errors.NoDocumentFound;
        }

        console.log(documents);

        sp.web.lists
          .getByTitle('DxpDokument')
          .items.filter(this.childrenFilter(documents))
          .get()
          .then((children: IDocumentModel[]) => {
            const documentList: IDocumentListItem[] = [];

            for (const parent of documents) {
              const listItem: IDocumentListItem = this.convertDocumentToItem(parent);

              for (const child of children) {
                if (child.DxpParentId === listItem.ID) {
                  listItem.children.push(child);
                }
              }

              documentList.push(listItem);
            }

            this.setState({ documents: documentList, selectedBook: item.key, error: undefined });
          });
      });
  }
  // Not necessary but maybe this mapper function will be needed
  private changeView(documentId: number): void {
    // sessionStorage.setItem('documentId', documentId.toString());

    this.props.changeView('documentView', documentId);
  }
}

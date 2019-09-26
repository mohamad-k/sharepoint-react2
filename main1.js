import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, Selection, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { IDragDropEvents, IDragDropContext } from 'office-ui-fabric-react/lib/utilities/dragdrop/interfaces';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { IconButton } from 'office-ui-fabric-react/lib/Button';
import { sp, SPBatch } from '@pnp/sp';

import IBookModel from '../../models/IBookModel';
import BookDialog from '../BookDialog';
import IDocumentModel from '../../models/IDocumentModel';
import styles from '../Docxpert.module.scss';

const columns: IColumn[] = [
  { key: 'id', name: 'ID', fieldName: 'ID', minWidth: 10, maxWidth: 20 },
  { key: 'title', name: 'Buch', fieldName: 'Title', minWidth: 100 },
  { key: 'action', name: 'Aktion', fieldName: undefined, minWidth: 100 }
];

let _draggedItem: any = undefined;
let _draggedIndex: number = -1;

export interface IBookOverviewProps {}

export interface IBookOverviewState {
  books: IBookModel[];
  isModalOpen: boolean;
  showDialog: boolean;
  editDialog: boolean;
  bookToUpdate: IBookModel[];
  selectedBook: IBookModel;
}

export default class BookSettingsView extends React.Component<IBookOverviewProps, IBookOverviewState> {
  public state: IBookOverviewState = {
    books: [],
    isModalOpen: false,
    showDialog: false,
    editDialog: false,
    bookToUpdate: [],
    selectedBook: undefined
  };

  constructor(props: {}) {
    super(props);

    this._editBook = this._editBook.bind(this);
    this._closeDialog = this._closeDialog.bind(this);
    this._showDialog = this._showDialog.bind(this);
    this._deleteBook = this._deleteBook.bind(this);
    this.onRenderItemColumn = this.onRenderItemColumn.bind(this);
    this._fetchBooks = this._fetchBooks.bind(this);
    this.renderBookDialog = this.renderBookDialog.bind(this);
    this._getDragDropEvents = this._getDragDropEvents.bind(this);
    this._insertBeforeItem = this._insertBeforeItem.bind(this);
    this._selection = new Selection();
  }

  public componentDidMount(): void {
    this._fetchBooks();
  }

  public render(): JSX.Element {
    return (
      <div className={styles.bookSettingsView}>
        <div className={styles.addButtonDiv}>
          <IconButton className={styles.addButton} iconProps={{ iconName: 'CirclePlus' }} onClick={this._showDialog} />
        </div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            dragDropEvents={this._getDragDropEvents()}
            items={this.state.books}
            columns={columns}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            layoutMode={DetailsListLayoutMode.justified}
            onRenderItemColumn={this.onRenderItemColumn}
            selectionMode={SelectionMode.multiple}
          />
        </MarqueeSelection>
        {this.renderBookDialog()}
      </div>
    );
  }

  private onRenderItemColumn(book: IBookModel, index: number, column: IColumn): JSX.Element {
    switch (column.key) {
      case 'title':
        return <span>{book.Title}</span>;
      case 'action':
        return (
          <div>
            <IconButton onClick={() => this._editBook(index)} iconProps={{ iconName: 'Edit' }} />
            <IconButton onClick={() => this._deleteBook(book.ID)} iconProps={{ iconName: 'Delete' }} />
          </div>
        );
      default:
        return <span>{book[column.fieldName]}</span>;
    }
  }
  private _selection: Selection;

  private _deleteBook(bookID: number): void {
    const confirmMessage: boolean = confirm('Sind Sie sicher, dass Sie das Buch und alle Documente darin löschen möchten?');
    if (confirmMessage) {
      sp.web.lists
        .getByTitle('DxpBuch')
        .items.getById(bookID)
        .delete()
        .then(() => {
          const batch: SPBatch = sp.createBatch();
          sp.web.lists
            .getByTitle('DxpDokument')
            .items.filter(`DxpBookId eq ${bookID}`)
            .get()
            .then((items: IDocumentModel[]) => {
              items.forEach((i: IDocumentModel) => {
                sp.web.lists
                  .getByTitle('DxpDokument')
                  .items.getById(i.ID)
                  .inBatch(batch)
                  .delete()
              });
              batch.execute().then(() => console.log('All deleted'));
            });
        })
        .then(() => {
          sp.web.lists
            .getByTitle('DxpBuch')
            .items.getAll()
            .then(books => {
              this.setState({ books });
            });
        });
    }
  }

  private _editBook(index: number): void {
    this.setState({ selectedBook: this.state.books[index], showDialog: true });
  }

  private _fetchBooks = (): void => {
    sp.web.lists
      .getByTitle('DxpBuch')
      .items.orderBy('Order1', true)
      .get()
      .then(books => {
        this.setState({ books, showDialog: false });
      });
  };
  private _showDialog = (): void => {
    this.setState({ showDialog: true });
  };
  private renderBookDialog = (): JSX.Element => {
    if (this.state.showDialog) {
      return <BookDialog isOpen={this.state.showDialog} close={this._closeDialog} reload={this._fetchBooks} book={this.state.selectedBook} />;
    }
  };
  private _closeDialog = (): void => {
    this.setState({ showDialog: false, selectedBook: undefined });
  };
  private _getDragDropEvents(): IDragDropEvents {
    return {
      canDrop: (dropContext?: IDragDropContext, dragContext?: IDragDropContext) => {
        return true;
      },
      canDrag: (item?: IBookModel) => {
        return true;
      },
      onDragEnter: (item?: IBookModel, event?: DragEvent) => {
        return 'dragEnter';
      }, // return string is the css classes that will be added to the entering element.
      onDragLeave: (item?: IBookModel, event?: DragEvent) => {
        return;
      },
      onDrop: (item?: IBookModel, event?: DragEvent) => {
        if (_draggedItem) {
          this._insertBeforeItem(item);
        }
      },
      onDragStart: (item?: IBookModel, itemIndex?: number, selectedItems?: IBookModel[], event?: MouseEvent) => {
        _draggedItem = item;
        _draggedIndex = itemIndex!;
      },
      onDragEnd: (item?: IBookModel, event?: DragEvent) => {
        _draggedItem = undefined;
        _draggedIndex = -1;
      }
    };
  }
  private _insertBeforeItem(item: IBookModel): void {
    const draggedItems: IBookModel[] = this._selection.isIndexSelected(_draggedIndex) ? this._selection.getSelection() : [_draggedItem];
    const books: IBookModel[] = this.state.books.filter(i => draggedItems.indexOf(i) === -1);
    let insertIndex: number = books.indexOf(item);

    if (insertIndex === -1) {
      insertIndex = 0;
    }
    books.splice(insertIndex, 0, ...draggedItems);
    this.setState({ books });

    const batch: SPBatch = sp.web.createBatch();
    const list = sp.web.lists.getByTitle('DxpBuch');

    for (let i: number = 0; i < books.length; i++) {
      list.items
        .getById(books[i].ID)
        .inBatch(batch)
        .update({ Order1: i });
    }
    batch.execute();
    console.log('all updated');
  }
}

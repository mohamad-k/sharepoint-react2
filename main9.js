import * as React from 'react';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { sp } from '@pnp/sp';
import { DetailsList, CheckboxVisibility, IColumn, IGroup } from 'office-ui-fabric-react/lib/DetailsList';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';
import { Link } from 'office-ui-fabric-react/lib/Link';

import IDocumentModel from '../models/IDocumentModel';
import IBookModel from '../models/IBookModel';
import IDocumentListItem from '../models/IDocumentListItem';
import { stringIsNullOrEmpty } from '@pnp/common';
import { Errors } from '../ErrorEnum';

export interface ISearchViewProps {
  documents: IDocumentModel[];
  error: string;
  changeView: Function;
}

export interface ISearchViewState {
  books: IBookModel[];
  error: string;
  groups: IGroup[];
}
export default class SearchView extends React.Component<ISearchViewProps, ISearchViewState> {
  public state: ISearchViewState = {
    books: [],
    error: undefined,
    groups: []
  };
  constructor() {
    super();

    this.errorMessage = this.errorMessage.bind(this);
    this.renderItemColumn = this.renderItemColumn.bind(this);
    // this.fetchDocuments = this.fetchDocuments.bind(this);
  }

  public componentDidMount(): void {
    // console.log(this.props.documents)
    // this.fetchDocuments(this.props.documents, this.props.error);
    // let documents: IDocumentModel[] = this.props.documents;

    // const books = [] as { ID: number; Title: string; count: number }[];
    // console.log(this.props.documents)

    // for (let i = 0; i < documents.length; i++) {

    //   if (books.filter(book => book.ID === documents[i].DxpBook.ID).length === 0){
    //     books.push({
    //       count: i,
    //       Title: documents[i].DxpBook.Title,
    //       ID: documents[i].DxpBook.ID
    //     });
    //   }
    //     console.log(books)
    // }
    
  }
  public componentWillReceiveProps(nextProps: ISearchViewProps): void {
    // this.fetchDocuments(nextProps.documents, nextProps.error);
    // const books = [] as { ID: number; Title: string; count: number }[];
    const {groups} = this.state 
    
    for (let i = 0; i < nextProps.documents.length; i++) {
      if (groups.filter(book => book.key === nextProps.documents[i].DxpBook.ID.toString()).length === 0){
        groups.push({
          count: 1 + i,
          name: nextProps.documents[i].DxpBook.Title,
          key: nextProps.documents[i].DxpBook.ID.toString(),
          startIndex:0
        });
      }
    }
    console.log(nextProps);
    console.log(groups)
    if(!nextProps.documents.length){
      this.setState({error : Errors.NoDocumentFound})
    }
  }

  public render(): JSX.Element {
    // let {books} = this.state;
    let { documents } = this.props;
    // let docFiltered: IDocumentModel[];
    // let docSorted: IDocumentModel[];

    // const isGroup: IGroup[] =
    //   return (
    //     {
    //       key: 'name',
    //       name: book.Title,
    //       startIndex:0 ,
    //       count: docFiltered.length
    //     }
    //   );
    // });
    if (this.props.documents.length) {
      return (
        <div>
          {/* {isGroup} */}
          <DetailsList
            items={documents}
            columns={[{ key: 'name', name: 'name', fieldName: 'Title', minWidth: 600 }]}
            setKey="set"
            selectionPreservedOnEmptyClick={true}
            checkboxVisibility={CheckboxVisibility.hidden}
            onRenderItemColumn={this.renderItemColumn}
            groups={this.state.groups}
          />
        </div>
      );
    } else if (this.state.error) {
      return this.errorMessage();
    } else {
      return <Spinner label="Bitte warten" />;
    }
  }
  // private childrenFilter(): string {
  //   let filter: string = '';

  //   for (let i: number = 0; i < this.props.documents.length; i++) {
  //     filter += `Id eq ${this.props.documents[i].DxpBookId}`;

  //     if (i < this.props.documents.length - 1) {
  //       filter += ' or ';
  //     }
  //   }
  //   console.log(filter)
  //   return filter;
  // }
  private errorMessage(): JSX.Element {
    if (this.state.error) {
      console.log(this.state.error);
      return (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {this.state.error}
        </MessageBar>
      );
    }
  }
  // private fetchDocuments(filterBooks: IDocumentModel[], errorMessage: string): void {
  //   sp.web.lists
  //     .getByTitle('DxpBuch')
  //     .items.filter(this.childrenFilter())
  //     .get()
  //     .then(books => {
  //       if (books.length > 0) {
  //         this.setState({ books, error: errorMessage });
  //       }
  //     });
  // }
  private renderItemColumn(item: IDocumentListItem, index: number, column: IColumn): JSX.Element {
    switch (column.key) {
      case 'name':
        return (
          <span>
            <Link onClick={() => this.changeView(item.ID)}>{item.Title}</Link>
          </span>
        );
    }
  }
  private changeView(documentId: number): void {
    this.props.changeView('documentView', documentId);
  }
}

import * as React from 'react';
import { Dialog, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { sp } from '@pnp/sp';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';

import IBookModel from '../models/IBookModel';
import IContentTypeModel from '../models/IContentTypeModel';

export interface IBookDialogProps {
  isOpen: boolean;
  close: Function;
  book?: IBookModel;
  reload: Function;
}

export interface IBookDialogState {
  bookTitle: string;
  contentTypes: IContentTypeModel[];
  selectedContentType: string;
}
export default class BookDialog extends React.Component<IBookDialogProps, IBookDialogState> {
  public state: IBookDialogState = {
    bookTitle: '',
    contentTypes: [],
    selectedContentType: undefined
  };

  constructor() {
    super();

    this.changeContentType = this.changeContentType.bind(this);
    this.changeBookName = this.changeBookName.bind(this);
    this.saveBook = this.saveBook.bind(this);
  }
  public componentDidMount(): void {
    console.log(this.props.book);
    sp.web.lists
      .getByTitle('DxpDokument')
      .contentTypes.get()
      .then(contentTypes => {
        console.log(contentTypes);
        this.setState({ contentTypes });
      });

    if (this.props.book) {
      this.setState({
        bookTitle: this.props.book.Title,
        selectedContentType: this.props.book.DxpContentType
      });
    }
  }
  public componentWillReceiveProps(nextProps: IBookDialogProps): void {
    console.log(nextProps.book);
  }
  public render(): JSX.Element {
    return (
      <div>
        <Dialog
          title={this.props.book ? 'Update Buch' : 'Neues Buch'}
          isOpen={this.props.isOpen}
          onDismiss={() => {
            this.props.close();
          }}
        >
          <TextField label="Buchname:" onChanged={this.changeBookName} required={true} value={this.state.bookTitle} />
          <Dropdown
            required={true}
            label="Inhaltstyp:"
            selectedKey={this.state.selectedContentType}
            onChanged={this.changeContentType}
            options={this.state.contentTypes.map(type => ({ key: type.Id.StringValue, text: type.Name }))}
          />
          <DialogFooter>
            <PrimaryButton onClick={this.saveBook} text="Save" />
            <DefaultButton
              onClick={() => {
                this.props.close();
              }}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }
  private changeContentType = (item: { key: string; text: string }): void => {
    this.setState({ selectedContentType: item.key });
  };
  private changeBookName = (text: string): void => {
    this.setState({ bookTitle: text });
  };
  private saveBook = (): void => {
    if (this.props.book) {
      sp.web.lists
        .getByTitle('DxpBuch')
        .items.getById(this.props.book.ID)
        .update({
          Title: this.state.bookTitle,
          DxpContentType: this.state.selectedContentType
        })
        .then(() => {
          this.props.reload();
          this.props.close();
        });
    } else {
      sp.web.lists
        .getByTitle('DxpBuch')
        .items.add({
          Title: this.state.bookTitle,
          DxpContentType: this.state.selectedContentType
        })
        .then(() => {
          this.props.reload();
          this.props.close();
        });
    }
  };
}

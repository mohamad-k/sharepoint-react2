import * as React from 'react';
import { sp } from '@pnp/sp';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Dropdown, IDropdown, DropdownMenuItemType, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { DatePicker, DayOfWeek, IDatePickerStrings, ICalendarFormatDateCallbacks } from 'office-ui-fabric-react/lib/DatePicker';
import { Label } from 'office-ui-fabric-react/lib/Label';

// import IBookModel from '../../models/IBookModel';
import IDocumentModel from '../models/IDocumentModel';
import IFieldModel from '../models/IFieldModel';

// import styles from '../Docxpert.module.scss';

export interface IContentTypeFormProps {
  document?: IDocumentModel;
  contentTypeId: string;
}

export interface IBookOverviewState {
  fields: IFieldModel[];
  options: object;
}

export default class ContentTypeForm extends React.Component<IContentTypeFormProps, {}> {
  public state: IBookOverviewState = {
    fields: [],
    options: {}
  };

  constructor() {
    super();

    this.renderDocument = this.renderDocument.bind(this);
  }

  public componentDidMount(): void {
    sp.web.contentTypes
      .getById(this.props.contentTypeId)
      .fields.get()
      .then(fields => {
        console.log(fields);

        const calls = [];

        for (let i = 0; i < fields.length; i++) {
          if (fields[i].TypeAsString === 'Lookup') {
            calls.push(this.fetchOptions(fields[i].LookupList, fields[i].LookupField, i));
          }
        }

        Promise.all(calls).then(results => {
          console.log(results);

          for (const result of results) {
            fields[result.index].options = result.result;
          }

          this.setState({ fields });
        });
      });
  }

  public render(): JSX.Element {
    return <div>{this.renderDocument()}</div>;
  }

  private renderDocument(): JSX.Element[] {
    if (this.props.document) {
      return Object.keys(this.props.document).map(key => (
        <div>
          <Label>{key}</Label>
          <span>{this.props.document[key]}</span>
        </div>
      ));
    } else {
      return this.state.fields.map(field => <div>{this.renderField(field)}</div>);
    }
  }

  private renderField(field: IFieldModel) {
    switch (field.TypeAsString) {
      case 'Text':
        return <TextField label={field.Title} />;
      case 'Lookup':
        // this.fetchOptions(field.LookupList, field.LookupField, field.Title);
        return (
          <Dropdown
            label={field.Title}
            // selectedKey={selectedItem ? selectedItem.key : undefined}
            // onChanged={this.changeState}
            placeHolder="Select an Option"
            options={field['options']}
          />
        );
      case 'DateTime':
        return (
          <DatePicker
            firstDayOfWeek={DayOfWeek.Monday}
            label={field.Title}
            showMonthPickerAsOverlay={true}
            formatDate={this.formateDate}
            // onAfterMenuDismiss={() => console.log('onAfterMenuDismiss called')}
          />
        );
    }
  }

  private fetchOptions(listId: string, display: string, index: number) {
    return sp.web.lists
      .getById(listId)
      .items.select('ID', display)
      .getAll()
      .then(result => {
        // console.log(result);
        // const options = this.state.options;
        // options[title] = result.map(option => ({ key: option.ID, text: option[display] }));
        // this.setState({ options });

        return { index, result: result.map(option => ({ key: option.ID, text: option[display] })) };
      });
  }

  private formateDate(date: Date) {
    return date.toLocaleDateString('de-DE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }
}

import * as React from 'react';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';

import { Views } from './ViewEnum';
import styles from './Docxpert.module.scss';
import IDocumentModel from '../models/IDocumentModel';

export interface INavbarProps {
  changeView: Function;
  search: Function;
}
export interface ISearchViewState {
  documents: IDocumentModel[];
}
export default class Navbar extends React.Component<INavbarProps, {}> {
  public state: ISearchViewState = {
    documents: []
  };
  constructor() {
    super();

    this.renderSearchBox = this.renderSearchBox.bind(this);
  }
  public render(): JSX.Element {
    const items: IContextualMenuItem[] = [
      {
        key: 'dashboard',
        name: 'Dashboard',
        iconProps: {
          iconName: 'Home'
        },
        onClick: e => {
          e.preventDefault();
          this.props.changeView(Views.Dashboard_VIEW);
        }
      },
      {
        key: 'tasks',
        name: 'Dokumente',
        iconProps: {
          iconName: 'CheckList'
        },
        onClick: e => {
          e.preventDefault();
          this.props.changeView(Views.TASK_VIEW);
        }
      },
      {
        key: 'books',
        name: 'Bücher',
        iconProps: {
          iconName: 'Dictionary'
        },
        onClick: e => {
          e.preventDefault();
          this.props.changeView(Views.BOOK_VIEW);
        }
      }
    ];

    const farItems: IContextualMenuItem[] = [
      {
        key: 'search',
        name: 'Suche',
        onRender: this.renderSearchBox
      },
      {
        key: 'settings',
        iconProps: {
          iconName: 'MoreVertical'
        },
        items: [
          {
            key: 'admBooks',
            name: 'Bücher',
            onClick: e => {
              e.preventDefault();
              this.props.changeView(Views.ADM_BOOK_VIEW);
            }
          }
        ],
        className: styles.smallBtn,
        onClick: e => {
          e.preventDefault();
        }
      }
    ];

    return <CommandBar className={styles.navbar} elipisisAriaLabel="More options" items={items} farItems={farItems} />;
  }

  private renderSearchBox(): JSX.Element {
    return (
      <SearchBox
        placeholder="Suche"
        value=""
        onSearch={value => {
          this.props.search(value);
        }}
      />
    );
  }
}

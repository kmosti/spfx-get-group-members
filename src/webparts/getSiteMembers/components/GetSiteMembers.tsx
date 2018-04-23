import * as React from 'react';
import styles from './GetSiteMembers.module.scss';
import {
  IGetSiteMembersProps,
  IGetSiteMembersState,
  IGroup,
  IGroupMember
} from './IGetSiteMembersProps';
import { Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from "sp-pnp-js";
import SimpleTable from 'react-simple-table';
// Import React Table
import ReactTable from "react-table";
import "react-table/react-table.css";
import 'babel-polyfill';

export default class GetSiteMembers extends React.Component<IGetSiteMembersProps, IGetSiteMembersState> {

  constructor(props: IGetSiteMembersProps, state: IGetSiteMembersState) {
    super(props);

    this.state = {
      loading: true,
      error: "",
      showError: false,
      groupTitle: "",
      results: [{
        Id: 0,
        Title: ""
      }]
    };
  }
  public componentDidMount(): void {
    this._processTasks();
  }

  public componentDidUpdate( prevProps: IGetSiteMembersProps, prevState: IGetSiteMembersState): void {
    if( prevProps.siteGroup !== this.props.siteGroup ) {
      this._resetLoadingState();
      this._processTasks();
    }
  }

  private _resetLoadingState() {
    this.setState({
        loading: true,
        error: "",
        showError: false
    });
}

  private _processTasks() {
    if(Number(this.props.siteGroup) ) {
      pnp.sp.web.siteGroups.getById(this.props.siteGroup).users.get().then( res => {
        let groupMembers: IGroupMember[] = res.map(person => ({ Title: person.Title, email: person.Email }));
        this.setState({
          loading: false,
          results: groupMembers
        });
      }).catch( err => {
        this.setState({
          loading: false,
          error: JSON.stringify(err)
        });
      });
    } else {
      this.setState({
        loading: false,
        error: "No group has been selected, please select group from the property pane dropdown"
      });
    }

  }

  public render(): React.ReactElement<IGetSiteMembersProps> {
    let view = <Spinner size={SpinnerSize.large} label="Loading" />;
    if (!this.state.loading && this.state.results) {
      view = <ReactTable
              data={this.state.results}
              columns={[{
                Header: this.props.groupTitle,
                columns: [{
                  Header: "Name",
                  accessor: "Title"
                },{
                  Header: "Email",
                  accessor: "email"
                }
              ]}
              ]}
              defaultPageSize={5}
              noDataText={"There are no members in this group"}
          />;
    }
    if (this.state.error !== "") {
      return (
        <MessageBar messageBarType={MessageBarType.error} className={styles.error}>
            <span>There was an error</span>
            {
                (() => {
                    if (this.state.showError) {
                        return (
                            <div>
                                <p>
                                    <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronUp`} aria-hidden="true"></i> Hide error message</a>
                                </p>
                                <p className="ms-font-m">{this.state.error}</p>
                            </div>
                        );
                    } else {
                        return (
                            <p>
                                <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronDown`} aria-hidden="true"></i> Show error message</a>
                            </p>
                        );
                    }
                })()
            }
        </MessageBar>
      );
    }
    return (
    <div className={ styles.getSiteMembers }>
        {view}
      </div>
    );
  }

  private _toggleError() {
    this.setState({
        showError: !this.state.showError
    });
  }
}

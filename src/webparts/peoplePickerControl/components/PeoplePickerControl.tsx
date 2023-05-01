import * as React from 'react';
//import styles from './PeoplePickerControl.module.scss';
import { IPeoplePickerControlProps } from './IPeoplePickerControlProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { IPeoplePickerControlState } from './IPeoplePickerControlState';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PrimaryButton } from 'office-ui-fabric-react';
export default class PeoplePickerControl extends React.Component<IPeoplePickerControlProps, IPeoplePickerControlState> {
  constructor(props: IPeoplePickerControlProps) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {
      Users: []
    }
  }

  onSaveData = () => {
    const { Users } = this.state;
    console.log(Users);
    const lstUserId: any[] = [];
    Users.forEach((element: { id: any; }) => {
      lstUserId.push(element.id);
    });
    console.log(lstUserId);
    sp.web.lists.getByTitle("Test2").items.add({
      //UserId: this.state.UserId      
      //UserId: {'results': Users.map((user: { id: any; })=> {user.id})}
      UserId: { 'results': lstUserId }
    })
      .then((response) => {
        console.log(`Item ${response.data.ID} created successfully.`);
      })
      .catch((error) => {
        console.log("Error", error);
      })
  }

  _getPeoplePickerItems = (items: any[]) => {
    console.log(items);
    this.setState({ Users: items });
  }

  public render(): React.ReactElement<IPeoplePickerControlProps> {
    return (
      <div> <PeoplePicker
        // context={this.props.context}
        context={this.props.context}
        personSelectionLimit={5}
        // defaultSelectedUsers={this.state.EmployeeName===""?[]:this.state.EmployeeName}
        required={false}
        onChange={this._getPeoplePickerItems}
        //defaultSelectedUsers={this.state.User.slice(0,5)}
        //defaultSelectedUsers={[this.state.Users?this.state.Users:""]}
        showHiddenInUI={false}
        principalTypes={[PrincipalType.User]}
        resolveDelay={1000}
        ensureUser={true}
      />
        <div style={{ textAlign: "center" }}>
          <PrimaryButton text="Create" onClick={() => this.onSaveData()} /></div>
      </div>
    );
  }
}

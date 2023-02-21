import * as React from 'react';
import { IArfDashboardProps } from './IArfDashboardProps';
import './ArfDashboardStyle.scss';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import Swal from 'sweetalert2';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Items } from '@pnp/sp/items';
export interface IHellospfxState {
  items
  : [
    {
      "Id": "",
      "Question": "",
      "Answer": "",
      "Date": "",
      "Department": "",
      "PrimaryEmailID": "",
      "SecondaryEmailID": ""
    }],
  title: string,
  users: any[],
  showMessageBar: false,
  ItemIds: any[],
  DepartmentState: [],
  // selectFrames: [],
  // selected_Id: '',
  // selected_Department: '',
  Department: '',
  // selectedItems: []
}
let listItemId: number = 0
export default class ArfDashboard extends React.Component<IArfDashboardProps, IHellospfxState> {
  // disabled: Boolean;
  constructor(props) {
    super(props);
    sp.setup({
      spfxContext: this.props.context
    });
    this.state = {

      showMessageBar: false,
      items: [
        {
          "Id": "",
          "Question": "",
          "Answer": "",
          "Date": "",
          "Department": "",
          "PrimaryEmailID": "",
          "SecondaryEmailID": ""
        }
      ],
      title: '',
      users: [],
      // showMessageBar: false,
      Department: '',
      DepartmentState: [],
      ItemIds: [],
      // selectFrames: [],
      // selected_Id: '',
      // selected_Department: '',
      // selectedItems: []
    };
  }

  public render(): React.ReactElement<IArfDashboardProps> {
    return (
      <div className="ARFDashboard_OrderKey_Maindiv1"  >
        <div id='ARFDashboard_MainDiv' >
          <div className='ARFDashboard_Head'>Dashboard Answer</div>
          <button className='ARFDashboard_show_Button' onClick={() => { this.openmodel_get(); }} >Show Dept</button>
          <button className='ARFDashboard_Add_Button' onClick={(_e) => this.openmodel_create()}>Add Dept</button>
          <div className="ARFDashboard_container1" id='ARFDashboard_container1_ID' >
            <ul className='ARFDashboard_ul'>
              {this.state.items.length > 0 ? this.state.items.map((items, key) => {
                return (
                  <form className='ARFDashboard_Form1' key={key}  >
                    <div className='ARFDashboard_label_div' >
                      <li className='ARFDashboard_li' >
                        <label className='ARFDashboard_label1' id='Questions'>
                          {items.Question}
                        </label>
                      </li>
                    </div>
                    <div className='ARFDashboard_Input_Button' >
                      <textarea className='ARFDashboard_text1' placeholder="Type Answer" typeof='textarea' id={items.Id} defaultValue={items.Answer} />
                      <select className='ARFDashboard_Select_Button' name="Department" defaultValue={items.Department} onChange={(e) => { this.getbyItem(items.Id, e); this.handleChange(e, key); }}  >
                        <option className='ARFDashboard_Select_Button_options' >Select Department</option>
                        {this.state.DepartmentState.map((items1: any, key1) => {
                          let selectedVal = false;
                          if (items1.Department == items.Department) {
                            selectedVal = true;
                          }
                          return (
                            <option className='ARFDashboard_Select_Button_options' key={key1} selected={selectedVal} defaultValue={items1.Department} >{items1.Department}</option>
                            // id="Department_item"
                          )
                        })}
                      </select>
                      <div defaultValue={items.PrimaryEmailID} id="Primary_ID" className='ARFDashboard_Select_Button_peoplepicker' >
                        <PeoplePicker

                          //  onChange={(e) =>{this.getbyItemPrimary(items.Id, e, key); }}
                          context={this.props.context}
                          placeholder="PrimaryEmailID"
                          personSelectionLimit={1}
                          showtooltip={true}
                          required={true}
                          showHiddenInUI={false}
                          ensureUser={true}
                          principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                          resolveDelay={1000}
                          defaultSelectedUsers={[items.PrimaryEmailID]}
                          disabled={!items.Department}
                          onChange={(e) => { this.onchangedTitle; this._getPeoplePickerItems; this.getbyItemPrimary(items.Id, e, key); }}
                        >{items.PrimaryEmailID}</PeoplePicker>
                      </div>
                      <div defaultValue={items.SecondaryEmailID} className='ARFDashboard_Select_Button_peoplepicker' id="Secondary_ID" >
                        <PeoplePicker
                          // onChange={(e) => { this.getbyItemSecondary(items.Id, e); }}
                          context={this.props.context}
                          placeholder="SecondaryEmailID"
                          personSelectionLimit={1}
                          showtooltip={true}
                          required={true}
                          disabled={!items.Department}
                          onChange={(e) => { this.onchangedTitle; this._getPeoplePickerItems; this.getbyItemSecondary(items.Id, e, key); }}
                          showHiddenInUI={false}
                          ensureUser={true}
                          principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                          // 
                          resolveDelay={1000}
                          defaultSelectedUsers={[items.SecondaryEmailID]}
                        >{items.SecondaryEmailID}</PeoplePicker>
                      </div>
                    </div>
                  </form>
                );
              }) :
                <div id="hidedata" className='ARFDashboard_Data'>Data Not Found</div>
              }
            </ul>
          </div>
          <button className='ARFDashboard_Add_Button_save_update' onClick={() => this.SaveItem()} >Save</button>
        </div>

        {/* Edit Department Model */}

        <div id="Department_Edit" className="ARFDashboard_modal_Edit"  >
          <form className="ARFDashboard_modal-content_ADD" >
            <span className="ARFDashboard_close" onClick={(_e) => this.closemodel_edit()} title="Close Modal">×</span>
            <div className="ARFDashboard_container1">
              <h1 className="ARFDashboard_header">Edit Departments</h1>
              <form className='ARFDashboard_Form1'>
                <div className='ARFDashboard_Input_Button_edit'>
                  <label className='ARFDashboard_label1'>Department</label>
                  <input className='ARFDashboard_text1_Edit' id="Department_EditItem" />
                  <button type="button" className="ARFDashboard_Add_Button1" onClick={() => this.updateItem()}>Save</button>
                </div>
              </form>
            </div>
          </form>
        </div>

        {/* Create Department Model */}

        <div id="Department_Create" className="ARFDashboard_modal" >
          <form className="ARFDashboard_modal-content_ADD" id="myForm" >
            <span className="ARFDashboard_close" onClick={(_e) => this.closemodel_create()} title="Close Modal">×</span>
            <div className="ARFDashboard_container1">
              <h1 className="ARFDashboard_header">Add Departments</h1>
              <form className='ARFDashboard_Form1'>
                <div className='ARFDashboard_Input_Button_create'>
                  <label className='ARFDashboard_label1'>Department</label>
                  <input className='ARFDashboard_text1_Create' id="Department" />
                  <button type="reset" className="ARFDashboard_Add_Button1" onClick={() => this.createItem()}  >Save</button>
                </div>
              </form>
            </div>
          </form>
        </div>

        {/* Get Department Model */}

        <div id="Department_get" className="ARFDashboard_modal">
          <form className="ARFDashboard_modal-content" >
            <span className="ARFDashboard_close" onClick={(_e) => this.closemodel_get()} title="Close Modal">×</span>
            <table className='ARFDashboard_Table' >
              <tr>
                <th className='ARFDashboard_Table_th'>
                  <span className='ARFDashboard_Table_span'>All Departments</span>
                </th>
                <th className='ARFDashboard_Table_th_1'><span className='ARFDashboard_Table_span'>Actions</span></th>
              </tr>
              {this.state.DepartmentState.map((items: any, key) => {
                return (
                  <tr className='ARFDashboard_Table_tr' key={key} >
                    <td className='ARFDashboard_Table_td' id='Department'>{items.Department}</td>
                    <td className='ARFDashboard_Table_td'>
                      <div className='ARFDashboard_button_main'>
                        <button className='ARFDashboard_Add_Button2' onClick={() => this.EditItem_Department(items.Id, items.Department)}>Edit</button>
                        <button className='ARFDashboard_Add_Button3' onClick={() => this.delete_Department(items.Id)} >Delete</button>
                      </div>
                    </td>
                  </tr>
                );
              }
              )

              }
            </table>
          </form>
        </div>
      </div>
    );

  }

  // Delete Departments

  public delete_Department(SelectedId) {
    debugger;
    if (SelectedId > 0) {
      this.props.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Department_RFQ')/items(${SelectedId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        })
        .then((response: SPHttpClientResponse): void => {
          if (response.ok) {
            // debugger;
            // alert(`Item ID: ${SelectedId} deleted successfully!`);
            Swal.fire({
              title: "Good job!",
              text: 'Delete successfully',
              type: "success",
              confirmButtonText: "ok",
              confirmButtonColor: "#bb4107",
            })

            this.getAllItems_Department();
          }

          else {
            alert(`Something went wrong!`);
            console.log(response.json());
          }
        }).then((newit: any): void => {
          console.log(newit);
        });
    }
    else {
      alert(`Please enter a valid item id.`);
    }
    event.preventDefault();
  }

  public componentDidMount(): void {
    this.PxHeight();
    this.getAllItems();
    this.getAllItems_Department();
  }

  // Get all Items from CustomerReq

  public getAllItems() {
    // debugger;
    var reactHandler = this;
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('CustomerREQ')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "odata-version": "",
        },
      })
      .then((response: SPHttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((itemobj: any): void => {
        // debugger;
        reactHandler.setState({ items: itemobj.value.filter(i => i.Answer == null || i.Answer == undefined || i.Answer == "") });
        console.log("items:" + itemobj);
      },
        (error: any): void => {
          console.log("Errors:" + error);
        }
      );
    // }

  }

  // Get Department

  public getAllItems_Department() {
    // debugger;
    var reactHandler = this;
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Department_RFQ')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "odata-version": "",
        },
      })
      .then((response: SPHttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((item: any): void => {
        reactHandler.setState({ DepartmentState: item.value });
        console.log("items:" + item);
      },
        (error: any): void => {
          console.log("Errors:" + error);
        }
      );
  }
  // Get openmodel

  public openmodel_get() {
    document.getElementById('Department_get').style.display = "block";
    this.getAllItems_Department();
  }
  // Create openmodel

  public openmodel_create() {
    document.getElementById('Department_Create').style.display = "block"
  }
  // Get closemodel

  public closemodel_get() {
    // debugger;
    document.getElementById('Department_get').style.display = "none";
    this.getAllItems();
  }
  // Create closemodel

  public closemodel_create() {
    // debugger;
    document.getElementById('Department_Create').style.display = "none"
    this.getAllItems();
  }
  // Edit model open

  public openmodel_edit() {
    document.getElementById("Department_Edit").style.display = "block";
  }
  // Edit closemodel
  public closemodel_edit() {
    document.getElementById('Department_Edit').style.display = "none"
  }

  // Create Department
  public createItem() {
    debugger;
    let Depatment_item = (document.getElementById("Department") as HTMLInputElement).value;
    if (document.getElementById("Department")["value"] != "") {
      console.log("Value of question", Depatment_item);
      const body: string = JSON.stringify({
        Department: Depatment_item
      });
      console.log("url", this.props.siteUrl)
      this.props.spHttpClient.post(
        `${this.props.siteUrl}/_api/web/lists/getbytitle('Department_RFQ')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=nometadata",
            "odata-version": "",
          },
          body: body,
        }
      )
        .then((response: SPHttpClientResponse): Promise<any> => {
          return response.json();
        })
        .then((newitem: any): void => {
          console.log(newitem);
          Swal.fire({
            title: "Good job!",
            text: 'Data Saved in List Successfully',
            type: "success",
            confirmButtonText: "ok",
            confirmButtonColor: "#bb4107",
          })

        },
          (_error: any): void => {
            console.log("Errors");
          }
        );
    }
  }
  // Full Screen

  public PxHeight() {
    // debugger;
    document.getElementById("sp-appBar") != undefined ?
      document.getElementById("sp-appBar").style.display = "none"
      : '';
    let screenWidth; let screenHeight;
    let finValue = screen.width; //1900
    let finValue1 = 'Not Done';
    setInterval(() => {
      screenWidth = screen.width;//800
      screenHeight = screen.height;
      if (screenWidth == finValue) {
        finValue = screen.width;//1000
        // screenWidth
        if (finValue1 != "Done") {
          finValue1 = 'Done';

          var element = document.getElementById(
            "ARFDashboard_MainDiv"
          );
          var back = document.getElementById("ARFDashboard_OrderKey_imageback");
          back.style.height = element.clientHeight + 90 + "px";

          console.log(element.clientHeight);

        }
      }
      else if (screenWidth != finValue) {
        finValue = screen.width;//800
        var element = document.getElementById(
          "ARFDashboard_MainDiv"
        );
        var back = document.getElementById("ARFDashboard_OrderKey_imageback");
        back.style.height = element.clientHeight + 90 + "px";

        console.log(element.clientHeight);
      }

    }, 500);

  }
  // Get all list items

  public getbyItem(Ids: any, e) {
    // debugger;
    var value = this.state.items.filter(x => x.Id == Ids);
    let value2;
    value2 = value[0].Id;
    let test = this.state.ItemIds.filter(x => x == value2)
    if (!test[0]) {
      this.state.ItemIds.push({ "ID": Ids, "Value": e.target.value });
      // this.state.ItemIds.push({ "ID": Ids, "Value": e.target.value, Department: e.target.value, "PrimaryEmailID1": e.target.value, "SecondaryEmailID1": e.target.value });
      this.setState({
        ItemIds: this.state.ItemIds,
      })
      console.log(this.state.ItemIds)
    }
  }

  // onchange method for PrimaryEmailID

  public getbyItemPrimary(Ids: any, e, index) {
    if (e.length == 0) {
      this.state.ItemIds[index].PrimaryEmailID = ""
    }
    else{
    // if(User !=""){
    debugger;
    //  if(items !== null && items.length <0){
    var value = this.state.items.filter(x => x.Id == Ids );
    let value2
    value2 = value[0].Id;
    let test = this.state.ItemIds.filter(x => x == value2)
    if (!test[0]) {
      this.state.ItemIds[index] = { ...this.state.ItemIds[index], "ID": Ids, "PrimaryEmailID": e[0].secondaryText };
      // this.state.ItemIds.push({...this.state.ItemIds[index], "ID": Ids, "PrimaryEmailID": e[0].secondaryText });
      
      this.setState({
        ItemIds: this.state.ItemIds,
      })
    }
  }
    console.log(this.state.ItemIds, "itemssss")
    // }
    // }
  }
  // onchange method for SecondaryEmailID

  public getbyItemSecondary(Ids: any, e, index) {
    debugger;
    if (e.length == 0) {
      this.state.ItemIds[index].SecondaryEmailID = ""
    }
    else{
    var value = this.state.items.filter(x => x.Id == Ids  );
    let value2
    value2 = value[0].Id;
    let test = this.state.ItemIds.filter(x => x == value2)
    if (!test[0]) {
      this.state.ItemIds[index] = { ...this.state.ItemIds[index], "ID": Ids, "SecondaryEmailID": e[0].secondaryText };
      this.setState({
        ItemIds: this.state.ItemIds,
      })
    }
      console.log(this.state.ItemIds, "items id")
    }

  }

  // Save and update in the list

  public SaveItem() {
    debugger;
    if (this.state.ItemIds.length > 0) {
      //  try {
      //  if(items !== null && items.length > 0){ 
      this.state.ItemIds.forEach((element , index1, array) => {
        var t = String(element.ID)
        if (index1 == this.state.ItemIds.length - 1) {
          var Datetext_last = "updated_Last";
          console.log("date", Datetext_last);
          console.log(array);
        }
        else{
          var Datetext_last = "updated";
        }
        // var Datetext_last = "updated";
        //  if ( element.Value!= "") {
        console.log("items post", element)
        const body: string = JSON.stringify({
          //  Answer: document.getElementById(t)["value"],
          ID: element.ID,
          Department: element.Value,
          PrimaryEmailID: element.PrimaryEmailID,
          SecondaryEmailID: element.SecondaryEmailID,
           Date: Datetext_last,
        });
        this.props.spHttpClient
          .post(
            `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('CustomerREQ')/items(${element.ID})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "Content-type": "application/json;odata=nometadata",
                "odata-version": "",
                "IF-MATCH": "*",
                "X-HTTP-Method": "MERGE",
              },
              body: body,
            }
          )
          .then((response: SPHttpClientResponse): void => {
            if (response.ok) {
              Swal.fire({
                title: "Good job!",
                text: 'Data Saved in List Successfully',
                type: "success",
                confirmButtonText: "ok",
                confirmButtonColor: "#bb4107",
              })
              this.getAllItems_Department();
            }
            else {
              response.json().then((responseJSON) => {
                console.log(responseJSON);
                alert(
                  `Something went wrong! Check the error in the browser console.`
                );
              });
            }
          })
          .catch((error) => {
            console.log(error);
          });

        // }
        // else{
        //   alert(
        //     'something is wrong'
        //   )
        // }
        //  }
      }
      );
      // }
      // else{
      //   alert("PeoplePicker Data is not entered")
      // }
    }
    else {
      alert('Data is not Entered ')
    }
    // }
    // catch(error) {
    //   console.log(error);
    //   alert(
    //     'Enter the Data'
    //   )
    // };
  }


  // Get Dropdrown

  public Department_dropdown = () => {
    debugger;
    var reactHandler = this;
    this.props.spHttpClient.get(`${this.props.siteUrl}/_api/web/lists/getbytitle('Department_RFQ')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=nometadata",
          "odata-version": "",
        },
      })
      .then((response: SPHttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((item: any): void => {
        reactHandler.setState({ items: item.value });
        console.log("items:" + item);
      },

        (error: any): void => {
          console.log("Errors:" + error);
        }

      );
    this.getAllItems();
  }

  // Edit Department

  public EditItem_Department(Id: any, Department: string) {
    debugger;
    document.getElementById("Department_Edit")["style"]["display"] = "block";
    document.getElementById("Department_EditItem")['value'] = Department;
    listItemId = Id;
    event.preventDefault();
  }

  // Update Department

  public updateItem = (): void => {
    debugger;
    if (document.getElementById("Department_EditItem")["value"] != "") {
      const body: string = JSON.stringify({
        Department: document.getElementById("Department_EditItem")['value'],
      });
      this.props.spHttpClient
        .post(
          `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Department_RFQ')/items(${listItemId})`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              Accept: "application/json;odata=nometadata",
              "Content-type": "application/json;odata=nometadata",
              "odata-version": "",
              "IF-MATCH": "*",
              "X-HTTP-Method": "MERGE",
            },
            body: body,
          }
        )
        .then((response: SPHttpClientResponse): void => {
          if (response.ok) {
            Swal.fire({
              title: "Good job!",
              text: 'Edit Question successfully',
              type: "success",
              confirmButtonText: "ok",
              confirmButtonColor: "#bb4107",
            })
            this.getAllItems_Department();
          } else {
            response.json().then((responseJSON) => {
              console.log(responseJSON);

              alert(
                `Something went wrong! Check the error in the browser console.`
              );
            });

          }
        })
        .catch((error) => {
          console.log(error);
        });
    }
  };

  // PeoplePicker method

  public _getPeoplePickerItems(items: any[]) {
    debugger;
    let getSelectedUsers = [];
    //  if(items !== null && items.length <0){
    for (let item in items) {
      getSelectedUsers.push(items[item].id);
    }
    this.setState({ users: getSelectedUsers });
    //  }
  }

  public onchangedTitle(title: string) {
    this.setState({ title: title });
  }

  //  private async _createItem() {
  //    try {
  //     await sp.web.lists.getByTitle("Project Details").items.add({
  //       Title: this.state.title,
  //       ProjectMembersId: { results: this.state.users }
  //    });
  //   }

  //     // this.setState({
  //     //    message: "Item: " + this.state.title + " - created successfully!",
  //     //     showMessageBar: true,
  //       //  messageType: MessageBarType.success
  //     // });

  //   }
  //   catch (error) {
  //     // this.setState({
  //     //    message: "Item " + this.state.title + " creation failed with error: " + error,
  //     //    showMessageBar: true,
  //       //  messageType: MessageBarType.error
  //     // });
  //   }
  // }

  // Disable/Enable PeoplePicker method

  public handleChange(e, index) {
    debugger;
    let temp = this.state.items;
    temp[index].Department = e.target.value;
    this.setState({ ...this.state, items: temp, Department: e.target.value })
    // this.setState({ ...this.state, Department: e.target.value })

  }

}

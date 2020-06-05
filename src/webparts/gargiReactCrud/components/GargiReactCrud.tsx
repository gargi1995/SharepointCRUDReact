import * as pnp from 'sp-pnp-js';
import * as React from 'react';
import styles from './GargiReactCrud.module.scss';
import { IGargiReactCrudProps } from './IGargiReactCrudProps';
import { IGargiReactCrudState } from './IGargiReactCrudState';
import { escape } from '@microsoft/sp-lodash-subset';
import { IStackTokens, Stack } from 'office-ui-fabric-react/lib/Stack';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { sp } from "@pnp/sp/presets/all";
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { DateTimePicker, DateConvention, TimeConvention } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { DropDownList } from '@progress/kendo-react-dropdowns';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPHttpClient,SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISPList {
  EmpID: string;
  EmployeeName: string;
  Experience: string;
  Location: string;
  Country:string;
  DateOfBirth: Date;
  //User: string[];
  }
  const dropdownStyles: Partial<IDropdownStyles> = {
    dropdown: { width: 300 }};
  // const options: IDropdownOption[] = [  
  //   { key: 'india', text: 'India' },
  //   { key: 'australia', text: 'Australia' },
  //   { key: 'russia', text: 'Russia'},
  //   { key: 'netherland', text: 'Netherland' },
  // ];
  const stackTokens: IStackTokens = { childrenGap: 20 };

  var items1: IDropdownOption[]=[];


export default class GargiReactCrud extends React.Component<IGargiReactCrudProps,IGargiReactCrudState> {
  public constructor(props: IGargiReactCrudProps, state: IGargiReactCrudState){ 
    super(props); 
    this.state = { 
      items: [],
      Date:new Date(),
      user:[],
      value: undefined,
      selectedTerms: [],
      managedmetadata:[]
    };
    sp.setup({
      spfxContext: this.context
    });
    //this._getValues();
  }
  private async _getValues() {
    const item: any = await sp.web.lists.getByTitle("ListOne").items.getById(1).get();
    this.setState({
      Date: new Date(item.Date)
    });
  }
  public render(): React.ReactElement<IGargiReactCrudProps> {
    return (
      <div className={ styles.gargiReactCrud }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Sharepoint List Details</span>
              <p className={ styles.subTitle }>Use webparts to customize your sharepoint page.</p>
              


              <div>
                <input type="text" id="EmpID" placeholder="EmpID"/>
                <input type="text" id='EmployeeName' placeholder="EmployeeName"/>
                <input type="text" id="Experience" placeholder="Experience"/>
                <input type="text" id="Location" placeholder="location"/>
              </div>




              <div>
              <Stack tokens={stackTokens}>
              <Dropdown id="Country" placeholder="Select an option" label="Country" options={this.state.items} styles={dropdownStyles} onChanged={this._handleChange}
                      />
               </Stack>
              </div>
              <div>
              <TaxonomyPicker allowMultipleSelections={true}
                termsetNameOrID="State"
                panelTitle="Select Term"
                label="State"
                context={this.props.context}
                onChange={this.onTaxPickerChange}
                isTermSetSelectable={false} />
              </div>
              <div>
                
                
                <div>
                <DateTimePicker label="Date Of Birth:"
                dateConvention={DateConvention.Date}
                timeConvention={TimeConvention.Hours12}
                value={this.state.Date}
                onChange={(date: Date) => this.setState({ Date: date })} />

                </div>

                <PeoplePicker    
               context={this.props.context}    
               titleText="People Picker"    
               personSelectionLimit={3}    
               groupName={""} 
               showtooltip={true}    
               isRequired={true}    
               disabled={false}    
               ensureUser={true}    
               selectedItems={this._getPeoplePickerItems}    
               showHiddenInUI={false}    
               principalTypes={[PrincipalType.User]}    
               resolveDelay={1000} />
              </div>
              
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
                <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
                  <a href="#" className={`${styles.button}`} onClick={() => this.AddItem()}>
                  <span className={styles.label}>Create item</span>
                  </a>&nbsp;
                </div>
              </div>
              <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className='ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1'>
              <a href="#" className={`${styles.button}`} onClick={() => this.UpdateItem()}>
                <span className={styles.label}>Update item</span>
              </a>&nbsp;
              <a href="#" className={`${styles.button}`} onClick={() => this.DeleteItem()}>
                <span className={styles.label}>Delete item</span>
              </a>
              <a href="#" className={`${styles.button}`} onClick={() => this.getListData()}>
                <span className={styles.label}>Read Item</span>
              </a>
            </div>
          </div>
           </div>
          </div>
        </div>
      </div>





    );
  }
  /*private AddEventListeners() : void{
    document.getElementById('AddItem').addEventListener('click',()=>this.AddItem());
    document.getElementById('UpdateItem').addEventListener('click',()=>this.UpdateItem());
    document.getElementById('DeleteItem').addEventListener('click',()=>this.DeleteItem());
   }*/
   AddItem()
  {
    var e = document.getElementById("Country");
    // console.log('People ',this.state.user[0]);
    // var peoplepicarray = [];  
    // for (let i = 0; i < this.state.user.length; i++) {  
    //   peoplepicarray.push(this.state.user[i]["id"]);  
    // } 
    var managedmetad={
      'Label':this.state.managedmetadata[0].name,//'3'
      'TermGuid':this.state.managedmetadata[0].key,//'134d2279-41aa-475b-ae6b-e12cf26097fd'
      'WssId':-1
    };  
//var result = e.options[e.selectedIndex].text;

    pnp.sp.web.lists.getByTitle('ListOne').items.add({
    EmpID:document.getElementById('EmpID')["value"],
    EmployeeName : document.getElementById('EmployeeName')["value"],
    Experience : document.getElementById('Experience')["value"],
    Location:document.getElementById('Location')["value"],
    Country:document.getElementById('Country').textContent,
    DateOfBirth: this.state.Date,
    UserIdId: this.state.user[0]["id"],
    Province: managedmetad
    // UserId: this.state.user[0]["id"]
});
// const body: string = JSON.stringify({
//   'EmpID':document.getElementById('EmpID')["value"],
//     'EmployeeName' : document.getElementById('EmployeeName')["value"],
//     'Experience' : document.getElementById('Experience')["value"],
//     'Location':document.getElementById('Location')["value"],
//     'Country':document.getElementById('Country').textContent,
//     'State': managedmetad,
//     'DateOfBirth': this.state.Date
//     });
// this.props.context.spHttpClient.post(`${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('ListOne')/items`, SPHttpClient.configurations.v1, 
//       { 
//         headers: { 
//               'Accept': 'application/json', 
//               'Content-type': 'application/json', 
//               'odata-version': ''
//               }, 
//         body: body 
//       }) 
//       .then((response: SPHttpClientResponse): Promise<ISPList>=> { 
//       return response.json(); 
//           }) 
//           .then((item: ISPList): void => {
//             console.log('Check xml',item);
//           alert('Item has been successfully Saved ');
//           }, (error: any): void => { 
//             alert(`${error}`); 
//           }); 
alert("Record with Employee Name : "+ document.getElementById('EmployeeName')["value"] + " Added !");

}
UpdateItem() 
{
  var managedmetad={
    'Label':this.state.managedmetadata[0].name,//'3'
    'TermGuid':this.state.managedmetadata[0].key,//'134d2279-41aa-475b-ae6b-e12cf26097fd'
    'WssId':-1
  };  
//var id= this.getId(document.getElementById('EmpID')["value"]);
//console.log('Fetched id',id);
var id = (parseInt(document.getElementById('EmpID')["value"])+1);
pnp.sp.web.lists.getByTitle("ListOne").items.getById(id).update({
  EmployeeName : document.getElementById('EmployeeName')["value"],
  Experience : document.getElementById('Experience')["value"],
  Location:document.getElementById('Location')["value"],
  Country:document.getElementById('Country').textContent,
  DateOfBirth: this.state.Date,
  UserIdId: this.state.user[0]["id"],
  Province: managedmetad
});
alert("Record with Employee Name : "+ document.getElementById('EmployeeName')["value"] + " Updated !");
 
}
DeleteItem() 
{
pnp.sp.web.lists.getByTitle('ListOne').items.getById(document.getElementById('EmpID')["value"]).delete();
alert("Record with Employee ID : "+ document.getElementById('EmpID')["value"] + " Deleted !"); 
}
private _getListData(): Promise<ISPList[]> {
  return pnp.sp.web.lists.getByTitle('ListOne').items.get().then((response) => {
  
  return response;
  });
  
  }
   private getId(empid){
    var id=-1;
    pnp.sp.web.lists.getByTitle("ListOne").items.get().then((itm: any[]) => {
    
    itm.forEach((item)=>{
      if(item.EmpId.trim()==empid.trim()){
        id=item.ID;
        
      }
    })
    ;
    });
    return id; 
    
  }
  private getListData(): void {
    // get all the items from a list
    pnp.sp.web.lists.getByTitle("ListOne").items.get().then((itm: any[]) => {
    
    });
    this._getListData()
    .then((response) => {
    //this._renderList(response);
    }); 
    }
    private _renderList(items: ISPList[]): void {
      let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
      html += `<th>EmployeeId</th><th>EmployeeName</th><th>Experience</th><th>Location</th>`;
      items.forEach((item: ISPList) => {
      html += `
      <tr>
      <td>${item.EmpID}</td>
      <td>${item.EmployeeName}</td>
      <td>${item.Experience}</td>
      <td>${item.Location}</td>
      </tr>
      `;
      });
      html += `</table>`;
      // const listContainer: Element = this.domElement.querySelector('#spGetListItems');
      // listContainer.innerHTML = html;
      }
      public async componentDidMount(): Promise<void>
      {
        // get all the items from a sharepoint list
        var reacthandler=this;
        //var i=pnp.sp.web.lists.getByTitle("Employee Page");
        //console.log(i);
        console.log('Inside Mount function');
        pnp.sp.web.lists.getByTitle("ListOne").fields.getByInternalNameOrTitle("Country").get().then(function(data){
          data.Choices.forEach((element)=>{
            items1.push({key:element, text:element});
          })
          // for(var k in data.Choices){
          //   items.push({key:data.title, text:data.title});
          // }
          console.log('Items pushed',items1);
          reacthandler.setState({items:items1});
          console.log('Items in the state',this.state.items);
          
          return items1;
       });
      }
      private onTaxPickerChange=(terms : IPickerTerms):void=> {
        console.log("Terms", terms);
        this.setState({managedmetadata:terms});}

private _handleChange = (item: IDropdownOption): void => { this.setState({ value: item }); } 

private _getPeoplePickerItems=(itempp: any[]):void=> {
  console.log('Items:',itempp);
  this.setState({user:itempp});
}
  
}

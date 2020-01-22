import * as React from 'react';
import styles from './CardInsert.module.scss';
import {IPropsInsert} from '../../../classes/IProps';
import {IStateInsert} from '../../../classes/IState';
import {IServiceClass} from '../../../classes/IService';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import { TextField, MaskedTextField } from 'office-ui-fabric-react/lib/TextField';
import { TaxonomyPicker, IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
import { PeoplePicker } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/components/Button';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { SPHttpClient , SPHttpClientResponse } from '@microsoft/sp-http'
import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemPanel,
  AccordionItemButton,
} from 'react-accessible-accordion';
import * as _ from 'lodash';
// Demo styles, see 'Styles' section below for some notes on use.
import 'react-accessible-accordion/dist/fancy-example.css';
import { string, func } from 'prop-types';


export default class CardInsert extends React.Component<IPropsInsert, IStateInsert> {
  public serviceclass : IServiceClass = new IServiceClass();
  private options: any[] = [];
  private cityoptions: any[] = [];
  public constructor(props: IPropsInsert, state: IStateInsert){     
    super(props);
    this.handleTitle = this.handleTitle.bind(this);
    this.handleDesc = this.handleDesc.bind(this);
    this._onCheckboxChange = this._onCheckboxChange.bind(this);
    this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    this.createItem = this.createItem.bind(this);
    this.onTaxPickerChange = this.onTaxPickerChange.bind(this);
    this._getManager = this._getManager.bind(this);
    this.handleaddress = this.handleaddress.bind(this);
    this.handlepincode = this.handlepincode.bind(this);
    this.handlepermaddress = this.handlepermaddress.bind(this);
    this.handlepermpincode = this.handlepermpincode.bind(this);
    this.state = {
      selectedItems: [],
      name: "",
      description: "",
      address: "",
      permenantaddress:"",
      permenantpincode: "",
      pincode: "",
      dpselectedItem: undefined,
      dpselectedcity: undefined,
      dpselectedstate: undefined,
      dpselectedpermstate: undefined,
      dpselectedpermcity: undefined,
      disablepermaddress: false,
      dpcity: [],
      dpStates: [],
      dppermcity: [],
      termKey: undefined,
      dpselectedItems: [],
      disableAddressToggle: false,
      defaultaddresscheck: false,
      disableToggle: false,
      defaultChecked: false,
      pplPickerType:"",
      userIDs: [],
      userManagerIDs: [],
      usermanageremailid: [],
      hideDialog: true,
      status: "",
      isChecked: false,
      showPanel: false,
      required: "This is Required",
      onSubmission:false,
      termnCond:false,
      cityid: "",
      stateid: "",
      statepermid: "",
      citypermid: ""
    }

    this.serviceclass.getWelcomeMessageDetails(this.props.context, `/_api/web/lists/getbytitle('State')/Items`).then((ticketitems: any) => {                
      ticketitems.value.forEach(item => {
        this.options.push({
          key: item.Id,
          text: item.Title
        });
      });
      
      this.setState({    
        dpStates: this.options
      });
    });

    this.serviceclass.getWelcomeMessageDetails(this.props.context, `/_api/web/lists/getbytitle('City')/Items?$select=Title,State/Id,ID&$expand=State'`).then((ticketitems: any) => {                
      ticketitems.value.forEach(item => {
        this.cityoptions.push({
          key: item.Id,
          text: item.Title,
          stateid: item.State.Id
        });
      });
    });
  }
  public render(): React.ReactElement<IPropsInsert> {debugger;
    const { dpselectedItem, dpselectedItems } = this.state;
    const { name, description } = this.state;   
    const { dpselectedcity, dpcity} = this.state;
    const { dpselectedstate, dpStates, dpselectedpermstate, dpselectedpermcity} = this.state;
    sp.setup({
      spfxContext: this.props.context
    });
    return (   
         
      <div className={styles.cardInsert}>
        
        <div className={styles.container}>
          <div className={styles.row}>
            <div className="ms-Grid-col ms-u-sm12">
              <span className={styles.header}>
                {this.props.description}
              </span>
            </div>
          </div>
          <div className={styles.row}>
            <div className="ms-Grid-col ms-u-sm12">
              <Accordion allowMultipleExpanded={true}>
                <AccordionItem>
                  <AccordionItemHeading>
                        <AccordionItemButton>
                            General Details
                        </AccordionItemButton>
                    </AccordionItemHeading>
                    <AccordionItemPanel >
                      <div className={`ms-Grid-row ${styles.row}`}>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                            <label className={styles.msLabel}>Employee Name</label>             
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                            <TextField value={this.state.name} required={true} onChanged={this.handleTitle}
                              errorMessage={(this.state.name.length === 0 && this.state.onSubmission === true) ? this.state.required : ""}/>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Job Description</label>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                          <TextField multiline autoAdjustHeight value={this.state.description} onChanged={this.handleDesc}
                            />
                        </div>
                      </div>
                      <div className={`ms-Grid-row ${styles.row}`}>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Project Assigned To</label>                
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                          <TaxonomyPicker
                            allowMultipleSelections={false}
                            termsetNameOrID="b3c45494-51fb-4c16-b9af-21cb87a15419"
                            panelTitle="Select Assignment"
                            label=""
                            context={this.props.context}
                            onChange={this.onTaxPickerChange}
                            isTermSetSelectable={false}
                            />
                            <p className={(this.state.termKey === undefined && this.state.onSubmission === true)? styles.fontRed : styles.hideElement}>This is required</p>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Department</label><br/>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                          <Dropdown
                            placeHolder="Select an Option"
                            label=""
                            id="component"
                            selectedKey={dpselectedItem ? dpselectedItem.key : undefined}
                            ariaLabel="Basic dropdown example"
                            options={[
                              { key: 'Choice 1', text: 'Choice 1' },
                              { key: 'Choice 2', text: 'Choice 2' },
                              { key: 'Choice 3', text: 'Choice 3' }
                            ]}
                            onChanged={this._changeState}
                            onFocus={this._log('onFocus called')}
                            onBlur={this._log('onBlur called')}
                            />
                        </div>
                      </div>
                      <div className={`ms-Grid-row ${styles.row}`}>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>External Hiring?</label>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                          <Toggle
                            disabled={this.state.disableToggle}
                            checked={this.state.defaultChecked}
                            label=""
                            onAriaLabel="This toggle is checked. Press to uncheck."
                            offAriaLabel="This toggle is unchecked. Press to check."
                            onText="On"
                            offText="Off"
                            onChanged={(checked) =>this._changeSharing(checked)}
                            onFocus={() => console.log('onFocus called')}
                            onBlur={() => console.log('onBlur called')}         
                          />
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Reporting Manager</label>
                        </div>
                        <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                          <PeoplePicker
                            context={this.props.context}
                            titleText=" "
                            personSelectionLimit={1}
                            groupName={""} // Leave this blank in case you want to filter from all users
                            showtooltip={false}
                            isRequired={true}
                            disabled={false}
                            selectedItems={this._getManager}
                            errorMessage={(this.state.userManagerIDs.length === 0 && this.state.onSubmission === true) ? this.state.required : " "}
                            
                            />
                        </div>
                      </div>
                    </AccordionItemPanel>
                </AccordionItem>
                <AccordionItem>
                  <AccordionItemHeading>
                    <AccordionItemButton>
                        Contact Details
                    </AccordionItemButton>
                  </AccordionItemHeading>
                  <AccordionItemPanel >
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm12">
                        <span className="subheader">
                          Present Address
                        </span>
                      </div>
                    </div>
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Address</label>             
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <TextField multiline autoAdjustHeight value={this.state.address} onChanged={this.handleaddress} />
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                        <label className={styles.msLabel}>Pin Code</label>
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <MaskedTextField mask="999999" value={this.state.pincode} onChanged={this.handlepincode} />
                      </div>
                    </div>
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>State</label>             
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <Dropdown
                          placeHolder="Select State"                      
                          id="statename"
                          selectedKey={dpselectedstate ? dpselectedstate.key : undefined}
                          ariaLabel="Basic dropdown example"
                          options={this.state.dpStates}
                          onChanged={this._fetchcity}
                        />
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                        <label className={styles.msLabel}>City</label>
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <Dropdown
                          placeHolder="Select City"
                          id="countryname"
                          selectedKey={dpselectedcity ? dpselectedcity.key : undefined}
                          ariaLabel="Basic dropdown example"
                          options={this.state.dpcity}
                          onChanged={this._setcity}
                        />
                      
                      </div>
                    </div>
                  
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm12">
                        <span className="subheader">
                          Permenant Address
                        </span>
                      </div>
                    </div>
                    <div className={`ms-Grid-row ${styles.row}`}>                      
                      <div className="ms-Grid-col ms-sm12">
                        <Toggle
                              label="Is Permenant Address same as Present Address"
                              disabled={this.state.disableAddressToggle}
                              checked={this.state.defaultChecked}
                              onAriaLabel="This toggle is checked. Press to uncheck."
                              offAriaLabel="This toggle is unchecked. Press to check."
                              onText="Yes"
                              offText="No"
                              onChanged={(checked) =>this._changeAddresstype(checked)}       
                            />
                      </div>
                    </div>
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>Address</label>             
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <TextField multiline 
                          disabled={this.state.disablepermaddress}
                          autoAdjustHeight value={this.state.permenantaddress} onChanged={this.handlepermaddress} />
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                        <label className={styles.msLabel}>Pin Code</label>
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <MaskedTextField mask="999999" 
                        value={this.state.permenantpincode}                         
                        disabled={this.state.disablepermaddress} onChanged={this.handlepermpincode} />
                      </div>
                    </div>
                    <div className={`ms-Grid-row ${styles.row}`}>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                          <label className={styles.msLabel}>State</label>             
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <Dropdown
                          placeHolder="Select State"                      
                          id="statename"
                          selectedKey={dpselectedpermstate ? dpselectedpermstate.key : undefined}
                          ariaLabel="Basic dropdown example"
                          options={this.state.dpStates}
                          onChanged={this._fetchpermcity}                            
                          disabled={this.state.disablepermaddress}
                        />
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg2 block">
                        <label className={styles.msLabel}>City</label>
                      </div>
                      <div className="ms-Grid-col ms-sm6 ms-md3 ms-lg4 block">
                        <Dropdown
                          placeHolder="Select City"
                          id="countryname"
                          selectedKey={dpselectedpermcity ? dpselectedpermcity.key : undefined}
                          ariaLabel="Basic dropdown example"
                          options={this.state.dppermcity}
                          onChanged={this._setpermcity} 
                          disabled={this.state.disablepermaddress}
                        />
                      
                      </div>
                    </div>
                  
                  </AccordionItemPanel>
                </AccordionItem>
              </Accordion>
            </div>
          </div>
          <div className={`ms-Grid-row ${styles.row}`}>            
            <div className="ms-Grid-col ms-sm12">
              <Checkbox onChange={this._onCheckboxChange} className="f-left" ariaDescribedBy={'descriptionID'} />I have read and agree to the terms & condition
              <p className={(this.state.termnCond === false && this.state.onSubmission === true)? styles.fontRed : styles.hideElement}>Please check the Terms & Condition</p>
            </div>     
          </div>
          <div className={`ms-Grid-row ${styles.row}`}>                           
            <div className="ms-Grid-col ms-u-sm2 block">
              <PrimaryButton text="Create" onClick={() => { this.validateForm(); }} />
            </div>
            <div className="ms-Grid-col ms-u-sm2 block">
              <DefaultButton text="Cancel" onClick={() => { this.setState({}); }} />
            </div>
          </div>

          <Panel
              isOpen={this.state.showPanel}
              type={PanelType.smallFixedFar}
              onDismiss={this._onClosePanel}
              isFooterAtBottom={false}
              headerText="Are you sure you want to create site ?"
              closeButtonAriaLabel="Close"
              onRenderFooterContent={this._onRenderFooterContent}
          ><span>Please check the details filled and click on Confirm button to create site.</span>
          </Panel>
          
          <Dialog
              hidden={this.state.hideDialog}
              onDismiss={this._closeDialog}
              dialogContentProps={{
                type: DialogType.normal,
                title: 'Request Submitted Successfully',
                subText: "" }}
                modalProps={{
                titleAriaId: 'myLabelId',
                subtitleAriaId: 'mySubTextId',
                isBlocking: false,
                containerClassName: 'ms-dialogMainOverride'            
                }}>
              <div dangerouslySetInnerHTML={{__html:this.state.status}}/>    
            <DialogFooter>
            <PrimaryButton onClick={()=>this.gotoHomePage()} text="Okay" />
            </DialogFooter>
          </Dialog>
        </div>
      </div>      
    );
  }

  private onTaxPickerChange(terms : IPickerTerms) {
    this.setState({ termKey: terms[0].key.toString() });
    console.log("Terms", terms);
  }
  
  private _getManager(items: any[]) {
    this.state.userManagerIDs.length = 0;
    for (let item in items)
    {   
      this.state.userManagerIDs.push(items[item].id);
      console.log(items[item].id);
    }
  }
  
  private _onRenderFooterContent = (): JSX.Element => {
    return (
      <div>
        <PrimaryButton onClick={this.createItem} style={{ marginRight: '8px' }}>
          Confirm
        </PrimaryButton>
        <DefaultButton onClick={this._onClosePanel}>Cancel</DefaultButton>
      </div>
    );
  }
  
  private _log(str: string): () => void {
    return (): void => {
      console.log(str);
    };
  }
  
  private _onClosePanel = () => {
    this.setState({ showPanel: false });
  }
  
  private _onShowPanel = () => {
    this.setState({ showPanel: true });
  }
  
  private _changeSharing(checked:any):void{
    this.setState({defaultChecked: checked});
  }

  private _changeAddresstype(checked:any):void{
    this.setState({defaultChecked: checked});
    console.log(this.state.pincode);
    this.getcity(this.state.stateid.toString(), true);
    if(checked){ 
      this.setState({
        permenantaddress: this.state.address, 
        permenantpincode: this.state.pincode, 
        statepermid: this.state.stateid,
        dpselectedpermstate: this.state.dpselectedstate,
        dpselectedpermcity: this.state.dpselectedcity, disablepermaddress: checked   
      });
    }
    else{
      this.setState({
        permenantaddress: "", permenantpincode: "", statepermid: "", dpselectedpermstate: undefined,
        dpselectedpermcity: undefined, disablepermaddress: false});
    }
  }
  
  private _changeState = (item: IDropdownOption): void => {
    console.log('here is the things updating...' + item.key + ' ' + item.text + ' ' + item.selected);
    this.setState({ dpselectedItem: item });
    if(item.text == "Employee")
    {
      this.setState({defaultChecked: false});
      this.setState({disableToggle: true});     
    }
    else
    {
      this.setState({disableToggle:false});
    }
  }
  
  private handleTitle(value: string): void {
    return this.setState({
      name: value
    });
  }
  
  private handleDesc(value: string): void {
    return this.setState({
      description: value
    });
  }

  private handleaddress(value: string): void{
    return this.setState({
      address: value
    })
  }

  private handlepincode(value: string): void{
    return this.setState({
      pincode: value
    })
  }

  private handlepermaddress(value: string): void{
    return this.setState({
      permenantaddress: value
    })
  }

  private handlepermpincode(value: string): void{
    return this.setState({
      permenantpincode: value
    })
  }

  private _fetchcity = (item: IDropdownOption): void =>{
    this.setState({stateid: item.key.toString(), dpselectedstate: item});
    this.getcity(item.key.toString(), false);    
  }

  private _fetchpermcity = (item: IDropdownOption): void =>{
    this.setState({statepermid: item.key.toString(), dpselectedpermstate: item});
    this.getcity(item.key.toString(), true);    
  }

  private _setpermcity = (item: IDropdownOption): void =>{
    this.setState({citypermid: item.key.toString(), dpselectedpermcity: item});
  }

  private _setcity = (item: IDropdownOption): void =>{
    this.setState({cityid: item.key.toString(), dpselectedcity: item});
  }  

  private getcity(cityid : string, ispermcondition) : void{
    let seletedcityoptions = _.filter(this.cityoptions, (o) => { return o.stateid == cityid });
    console.log(seletedcityoptions);
    if(!ispermcondition){
      this.setState({    
        dpcity: seletedcityoptions
      });
    }

    else{
      this.setState({    
        dppermcity: seletedcityoptions
      });
    }
  }
  
  private _onCheckboxChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean): void {
    console.log(`The option has been changed to ${isChecked}.`);
    this.setState({termnCond: (isChecked)?true:false});
  }
  
  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
  
  private _showDialog = (status:string): void => {   
    this.setState({ hideDialog: false });
    this.setState({ status: status });
  }
  
  private validateForm():void{
    let allowCreate: boolean = true;
    this.setState({ onSubmission : true });
    
    if(this.state.name.length === 0)
    {
      allowCreate = false;
    }
    if(this.state.termKey === undefined)
    {
      allowCreate = false;
    }   
    
    if(allowCreate)
    {
      this._onShowPanel();
    }
    else
    {
      //do nothing
    } 
  }
  
  private gotoHomePage():void{
    window.location.replace(this.props.siteurl);
  }

  private createItem():void { debugger;
    if(this.state.userManagerIDs[0] != null){
      this.serviceclass.getCurrentUserId(this.props.context, this.state.userManagerIDs[0].toString()).then((ticketitems: any) => {    
        let repmanager = ticketitems;
        this._onClosePanel(); 
        
        console.log(this.state.termKey);
        
        let postUrl = `${this.props.siteurl}/_api/web/lists/getbytitle('Employee Registration')/items`;
        let headers : any = {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        };
        let itembody : any= {
          "__metadata": {"type": "SP.Data.Employee_x0020_RegistrationListItem"},
          "Title": this.state.name,
          "Description": this.state.description,
          "Department": this.state.dpselectedItem.key,
          "Projects": {
              __metadata: { "type": "SP.Taxonomy.TaxonomyFieldValue" },
              Label: "1",
              TermGuid: this.state.termKey,
              WssId: -1
          },
          "ReportingManagerId": repmanager
        }
        this.props.context.spHttpClient.post(postUrl,SPHttpClient.configurations.v1,
        {
          headers: headers,
          body: JSON.stringify(itembody)
        }).then((response: SPHttpClientResponse): void => {
          this._showDialog("New Employee Created");
        }, (error: any): void => {
          console.log(error);
        });
        
      });
    }    
  }
}

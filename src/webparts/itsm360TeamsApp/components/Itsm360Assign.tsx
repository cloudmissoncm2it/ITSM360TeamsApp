import * as React from 'react';
import { Modal,Alert,Button,Icon } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserDetails } from '../model/IUserDetails';

export interface IItsm360AssignProps{
    visible:boolean;
    sharepointservice:sharepointservice;
    selectedTicket:ITicketItem;
    ppcontext:WebPartContext;
}

export interface IItsm360AssignState{
    ismodalvisible:boolean;
    modalsave?:boolean;
    assignpersonid?:string;
    errorMessage?:boolean;
}

export class Itsm360Assign extends React.Component<IItsm360AssignProps,IItsm360AssignState>{

    constructor(props:IItsm360AssignProps){
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false,
            errorMessage:false
        };
    }

    public handleOk=(e)=>{
        const {assignpersonid}=this.state;
        this.setState({modalsave:true});
        if(typeof assignpersonid !="undefined" && assignpersonid!="0"){
            console.log("People is selected");
            this.props.sharepointservice.updateTicketAssign(this.props.selectedTicket.ID,assignpersonid).then((data)=>{
                this.setState({
                    modalsave:false,
                    ismodalvisible:false,
                    errorMessage:false
                });
            });
            
        }else{
            this.setState({
                errorMessage:true,
                modalsave:false
            });
        }
        
    }

    public handleCancel=(e)=>{
        this.setState({ismodalvisible:false});
    }

    public _getAssignePickerItems=(people: any[])=> {
        /**
         * structure of the people object
         * id: "i:0#.f|membership|tka_itsmcompany.net#ext#@cloudmission.net"
          imageInitials: "TK"
          imageUrl: "https://cloudmission.sharepoint.com/sites/ThirumalITSMDev/_layouts/15/userphoto.aspx?accountname=tka%40itsmcompany.net&size=M"
          optionalText: ""
          secondaryText: "tka@itsmcompany.net"
          tertiaryText: ""
          text: "Thirumal Kandari"
         */
        let assignedpersonid:IUserDetails[]=[];
        if(people.length>0){
            assignedpersonid=this.props.sharepointservice._lusers.filter(i=>i.Email==people[0].secondaryText);
        }
        this.setState({assignpersonid:assignedpersonid.length>0?assignedpersonid[0].ID:"0"});
    }

    public render():React.ReactElement<IItsm360AssignProps>{
        const ticketdesc=this.props.selectedTicket?this.props.selectedTicket.Title:"";
        const ticketid=this.props.selectedTicket?this.props.selectedTicket.ID:"";
        const presentAssigne=this.props.selectedTicket?this.props.selectedTicket.AssignedPerson:"";

        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={()=> this.setState({ismodalvisible:true})}>
                <Icon type="user-add" />
                    Assign/Re-Assign 
                </Button>
                <Modal title="Assign/Re-Assign"
                   visible={this.state.ismodalvisible}
                   onOk={this.handleOk}
                   onCancel={this.handleCancel}
                   okText="Update Assign"
                   confirmLoading={this.state.modalsave} 
                >
                    <div>
                        
                       {this.state.errorMessage?<Alert type="error" style={{marginBottom:"10px"}} closable message="No user selected, select the user before submitting."/>:""} 
                        <Alert type="info" showIcon style={{marginBottom:"10px"}} 
                            message={`Ticket Title: ${ticketdesc}`}
                        />
                        <Alert type="info" showIcon style={{marginBottom:"10px"}}
                            message={`Ticket Id: ${ticketid}`}
                        />
                        <Alert type="info" showIcon style={{marginBottom:"10px"}}
                            message={`Present Value: ${presentAssigne}`}
                        />
                        <div className="ant-form ant-form-horizontal">
                        <div className="ant-row ant-form-item" >
                    <div className="ant-col ant-form-item-label ant-col-xs-24 ant-col-sm-8">
                      <label>Assign/Re-Assign</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper ant-col-xs-24 ant-col-sm-16">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                        <PeoplePicker
                            context={this.props.ppcontext}
                            titleText=""
                            personSelectionLimit={1}
                            showtooltip={true}
                            isRequired={false}
                            disabled={false}
                            selectedItems={this._getAssignePickerItems}
                            showHiddenInUI={false}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={200} />
                        </span>
                      </div>
                    </div>
                  </div>
                        </div>
                    </div>
                </Modal>
            </div>
        );
    }
}
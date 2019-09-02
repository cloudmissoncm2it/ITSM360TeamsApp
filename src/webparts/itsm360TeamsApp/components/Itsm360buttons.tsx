import * as React from 'react';
import { Select, Button, Input, Icon, Tabs, Modal, Cascader } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITeam } from '../model/ITeam';
import { Istatus } from '../model/Istatus';
import { IUserDetails } from '../model/IUserDetails';
import {Itsm360Attachment} from './Itsm360Attachment';
import { ITicketItem } from '../model/ITicketItem';
import {Itsm360StatusUpdate}from './Itsm360StatusUpdate';
import {Itsm360Assign} from './Itsm360Assign';

const { Option } = Select;
const {TextArea}=Input;
const{TabPane}=Tabs;

export interface IItsm360Props {
  hasSelected: boolean;
  sharepointservice: sharepointservice;
  ppcontext:WebPartContext;
  teams:ITeam[];
  status:Istatus[];
  selectedTicket:ITicketItem;
  refreshticketsdata?:any;
}

export interface IItsm360State {
  modalvisible: boolean;
  modalsave: boolean;
  cascaderoptions?: any[];
  cascaderdefaultvalue?:string[];
  title?: string;
  requestor?:any[];
  origin?:string;
  servicegroup?:string;
  service?:string;
  subcategory?:string;
  impact?:string;
  urgency?:string;
  description?:string;
  teamid?:string;
  statusid?:string;
  assignedperson?:any[];
  internalnote?:string;
  isAttachmentVisible?:boolean;
}

export class Itsm360buttons extends React.Component<IItsm360Props, IItsm360State>{

  constructor(props: IItsm360Props) {
    super(props);
    this.state = {
      modalvisible: false,
      modalsave: false,
      requestor:[this.props.sharepointservice._currentuser.email],
      urgency:"Low",
      impact:"Low",
      isAttachmentVisible:false
    };
  }

  public componentDidMount() {
    this.props.sharepointservice.getlookupdatanew().then((data) => {
      this.setState({ cascaderoptions: data });
    });
  }

  public handleOk = (e) => {
    this.setState({ modalsave: true });
    const {title,requestor,origin,servicegroup,service,subcategory,impact,urgency,description,teamid,assignedperson,internalnote,statusid}=this.state;
    let requestorid:IUserDetails[]=[];
    if(typeof requestor !="undefined" && requestor.length>0){
      requestorid=this.props.sharepointservice._lusers.filter(i=>i.Email==requestor[0].secondaryText);
    }
    let assignedpersonid:IUserDetails[]=[];
    if(typeof assignedperson !="undefined" && assignedperson.length>0){
      assignedpersonid=this.props.sharepointservice._lusers.filter(i=>i.Email==assignedperson[0].secondaryText);
    }
    const newticket={
      Title:title,
      RequesterId:requestorid.length>0?requestorid[0].ID:null,
      Origin:origin,
      ServiceGroupsId:servicegroup,
      RelatedServicesId:service,
      RelatedCategoriesId:subcategory,
      Impact:impact,
      Urgency:urgency,
      Description:description,
      AssignedTeamId:teamid,
      AssignedPersonId:assignedpersonid.length>0?assignedpersonid[0].ID:null,
      Notes:internalnote,
      TicketsStatusId:statusid
    };

    console.log(newticket);

    this.props.sharepointservice.addITSMTicket(newticket).then((result)=>{
      console.log("post success: ", result);
      //this.props.refreshticketsdata();
      this.setState({
        modalvisible: false,
        modalsave:false
      });
      // this.setState({
      //   modalvisible: false,
      //   modalsave:false,
      //   title:"",
      //   requestor:[this.props.sharepointservice._currentuser.email],
      //   origin:"",
      //   impact:"Low",
      //   urgency:"Low",
      //   description:"",
      //   cascaderdefaultvalue:["0","0","0"]
      // });
    }).catch((error: any) => {
      console.log("Error: ", error);
    });

    
  }

  public handleCancel = (e) => {
    this.setState({
      modalvisible: false
    });
  }


  public showModal = () => {
    this.setState({
      modalvisible: true
    });
  }

  public titleChange = (e) => {
    this.setState({ title: e.currentTarget.value });
  }

  public descriptionChange = (e) => {
    this.setState({ description: e.currentTarget.value });
  }

  public internalNoteChange = (e) => {
    this.setState({ internalnote: e.currentTarget.value });
  }

  public _getPeoplePickerItems=(people: any[])=> {
    console.log(people);
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
    this.setState({requestor:people});
  }

  public _getAssignePickerItems=(people: any[])=> {
    console.log(people);
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
    this.setState({assignedperson:people});
  }

  public originChange = (value) => {
    this.setState({origin:value});
  }

  public teamChange = (value) => {
    this.setState({teamid:value});
  }

  public statusChange = (value) => {
    debugger;
    this.setState({statusid:value});
  }
  
  public cascaderChange = (e) => {
    console.log(e);
    if(e.length==3){
      this.setState({
        servicegroup:e[0],
        service:e[1],
        subcategory:e[2]
      });
    }else if(e.length==2){
      this.setState({
        servicegroup:e[0],
        service:e[1],
        subcategory:"-1"
      });
    } else if(e.length==1){
      this.setState({
        servicegroup:e[0],
        service:"-1",
        subcategory:"-1"
      });
    }else{
      this.setState({
        servicegroup:"-1",
        service:"-1",
        subcategory:"-1"
      });
    }
  }

  public urgencyChange=(value) => {
    this.setState({urgency:value});
  }

  public impactChange=(value) => {
    this.setState({impact:value});
  }

  public render(): React.ReactElement<IItsm360Props> {
    return (
      <div>
        <div className="gutter-box">
          <Itsm360StatusUpdate visible={this.props.hasSelected} sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket} status={this.props.status} />
          <Itsm360Assign visible={this.props.hasSelected} sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket} ppcontext={this.props.ppcontext} />
          <Itsm360Attachment visible={this.props.hasSelected} sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket} />
          <Button disabled={!this.props.hasSelected} >
            <Icon type="check-square" />
            Resolve
          </Button>
          <Button onClick={this.showModal}>
            <Icon type="plus-square" />
            New Ticket
          </Button>
        </div>
        <Modal title="New Ticket"
          visible={this.state.modalvisible}
          onOk={this.handleOk}
          onCancel={this.handleCancel}
          okText="Submit"
          confirmLoading={this.state.modalsave}
          destroyOnClose={true}
        >
          <div className="card-container">
            <Tabs type="card">
              <TabPane tab="Record" key="1">
                <div className="ant-form ant-form-vertical">
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Title</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <Input type="text" className="ant-input" placeholder="Ticket Title" onChange={this.titleChange} value={this.state.title} />
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Requestor</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <PeoplePicker
                            context={this.props.ppcontext}
                            titleText=""
                            personSelectionLimit={1}
                            showtooltip={true}
                            isRequired={false}
                            disabled={false}
                            selectedItems={this._getPeoplePickerItems}
                            showHiddenInUI={false}
                            principalTypes={[PrincipalType.User]}
                            resolveDelay={200} />
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item" >
                    <div className="ant-col ant-form-item-label">
                      <label>Origin</label>
                    </div>
                    <div className="ant-col ant-form-ioriginChangetem-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <Select placeholder="Please select Origin" style={{ width: "40%" }} onChange={this.originChange} value={this.state.origin}>
                            <Option value="Phone">Phone</Option>
                            <Option value="Email">Email</Option>
                            <Option value="Self-Service">Self-Service</Option>
                            <Option value="Walk-in">Walk-in</Option>
                            <Option value="Instant messaging">Instant messaging</Option>
                            <Option value="Client meeting">Client meeting</Option>
                            <Option value="Forwarded">Forwarded</Option>
                          </Select>
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Classification (Categories)</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <Cascader options={this.state.cascaderoptions} onChange={this.cascaderChange} placeholder="please select" defaultValue={this.state.cascaderdefaultvalue} />
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item" >
                    <div className="ant-col ant-form-item-label">
                      <label>Priority</label>
                    </div>
                    <div className="ant-col ant-form-ioriginChangetem-control-wrapper">
                      <div className="ant-form-item-control" style={{ marginBottom: "10px" }}>
                        <span className="ant-form-item-children" >
                          Impact
                  </span>
                        <span className="ant-form-item-children" style={{ marginLeft: "20px" }}>
                          <Select style={{ width: "40%" }} defaultValue="Low" onChange={this.impactChange} value={this.state.impact}>
                            <Option value="Low">Low</Option>
                            <Option value="Mid">Mid</Option>
                            <Option value="High">High</Option>
                          </Select>
                        </span>
                      </div>
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          Urgency
                  </span>
                        <span className="ant-form-item-children" style={{ marginLeft: "20px" }}>
                          <Select style={{ width: "40%" }} onChange={this.urgencyChange} value={this.state.urgency}>
                            <Option value="Low">Low</Option>
                            <Option value="Mid">Mid</Option>
                            <Option value="High">High</Option>
                          </Select>
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Description</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <TextArea placeholder="Summarize the request or issue" rows={4} onChange={this.descriptionChange} value={this.state.description} />
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </TabPane>
              <TabPane tab="Assign" key="2">
                <div className="ant-form ant-form-vertical">
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Assign Team</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <Select placeholder="Please select a team" style={{ width: "40%" }} onChange={this.teamChange}>
                            {this.props.teams.map((team: ITeam, index) => <Option value={team.ID} key={index}>{team.Title}</Option>)}
                          </Select>
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Assign Person</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
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
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Status</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <Select placeholder="select a status" style={{ width: "40%" }} onChange={this.statusChange}>
                            {this.props.status.map((stat: Istatus, index) => <Option value={stat.ID} key={index}>{stat.Title}</Option>)}
                          </Select>
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="ant-row ant-form-item">
                    <div className="ant-col ant-form-item-label">
                      <label>Internal Note</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                          <TextArea placeholder="Make a note" rows={3} onChange={this.internalNoteChange} value={this.state.internalnote} />
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
              </TabPane>
            </Tabs>
          </div>
          
        </Modal>
        
      </div>
    );
  }
}
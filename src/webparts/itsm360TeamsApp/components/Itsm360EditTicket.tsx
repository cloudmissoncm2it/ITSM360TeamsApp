import * as React from 'react';
import { Drawer, Button, Row, Col, Divider, Tabs, Form, Select, Input, List, Comment,Descriptions,Avatar } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import * as moment from 'moment';
import { ITeam } from '../model/ITeam';
import { Istatus } from '../model/Istatus';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserDetails } from '../model/IUserDetails';
import {Itsm360Classification} from './Itsm360Classification';
import {Itsm360InternalNotes} from './Itsm360InternalNotes';
import {Itsm360SubTasks}from './Itsm360SubTasks';
import {Itsm360Conversation} from './Itsm360Conversation';
import {Itsm360OrderDetails} from './Itsm360OrderDetails';
import { Itsm360Emails } from './Itsm360Emails';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
const { TabPane } = Tabs;
const { Option } = Select;
const { TextArea } = Input;


export interface IItsm360EditTicketProps {
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
    teams?: ITeam[];
    status?: Istatus[];
    ppcontext: WebPartContext;
    tictitle: string;
    refreshticketsdata?:any;
}

export interface IItsm360EditTicketState {
    isdrawervisible?: boolean;
    modalsave?: boolean;
    iserror?: boolean;
    errorMessage?: string;
    assignedperson?: any[];
    requester?: any[];
    notesdata?: any[];
    emails?: any[];
    newnote?: string;
    isStatusClosed?: boolean;
    newat?: string;
    newstatus?: string;
    closingComments?: string;
    loading?: boolean;
    ticketattachments?: any[];
    tickettitle?: string;
    ticketdescription?: string;
    ticketimpact?: string;
    ticketurgency?: string;
    servicegroup?: string;
    service?: string;
    subcategory?: string;
    classificationvalues?:any;
    internalnotesvalues?:any;
    notificationsummary?:any;
    orderdetails?:any[];
}

export class Itsm360EditTicket extends React.Component<IItsm360EditTicketProps, IItsm360EditTicketState>{

    private _sgtitle:string;
    private _setitle:string;
    private _scategorytitle:string;
    private _count=0;

    constructor(props: IItsm360EditTicketProps) {
        super(props);
        this.state = {
            isdrawervisible: false,
            modalsave: false,
            iserror: false,
            notesdata: [],
            emails:[],
            isStatusClosed: false,
            loading: false,
            ticketattachments: [],
            tickettitle: this.props.selectedTicket.Title,
            ticketimpact: "Low",
            ticketurgency: "Low",
            ticketdescription:"",
            orderdetails:[]
        };
    }

    public handleClose = (e) => {
        this.setState({ isdrawervisible: false });
    }

    public _getPeoplePickerItems = (people: any[]) => {
        this.setState({ assignedperson: people });
    }

    public _getrequesterpeopleItems = (people: any[]) => {
        this.setState({ requester: people });
    }

    public editTicketClick = (e) => {
        this.setState({ isdrawervisible: true });
        const { Status, ID } = this.props.selectedTicket;
        const spservice = this.props.sharepointservice;
        if (Status.indexOf("Closed") > -1) {
            this.setState({ isStatusClosed: true });
        }
        spservice.getTicketAttachment(ID).then((ticketattach) => {
            this.setState({ ticketattachments: ticketattach });
        });

        spservice.getTicketDetails(ID).then((ticketdata) => {
            ++this._count;
            this.setState({
                ticketdescription: ticketdata.Description,
                ticketurgency: ticketdata.Urgency,
                ticketimpact: ticketdata.Impact,
                notificationsummary:ticketdata.NotificationSummary
            });
            this._sgtitle=ticketdata.ServiceGroups.Title;
            this._setitle=ticketdata.RelatedServices.Title;
            this._scategorytitle=ticketdata.RelatedCategories.Title;
            const ciresults=ticketdata.RelatedCIs.results;
            const ras=ticketdata.RelatedAssets.results;
            let raas:any[]=[];
            ras.forEach((ra)=>{
                raas.push(ra.ID);
            });
            const xyz=JSON.parse(ticketdata.OrderDetails);
            debugger;
            this.setState({
                classificationvalues:{
                    servicegroupid:ticketdata.ServiceGroups.ID,
                    serviceid:ticketdata.RelatedServices.ID,
                    subcategoryid:ticketdata.RelatedCategories.ID,
                    relatedCiid:ciresults.length>0?ciresults[0].ID:"",
                    relatedCititle:ciresults.length>0?ciresults[0].Title:"",
                    servicegrouptitle:this._sgtitle,
                    servicetitle:this._setitle,
                    subcategorytitle:this._scategorytitle,
                    relatedassets:raas,
                },
                orderdetails:xyz
            });
        });
    }

    public postinternalnotes = () => {
        const { newnote, notesdata } = this.state;
        if (typeof newnote != "undefined") {
            const Currentusers: IUserDetails[] = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.sharepointservice._currentuser.email);
            const user: IUserDetails = Currentusers.length > 0 ? Currentusers[0] : null;
            const tnote = {
                TicketIDId: this.props.selectedTicket.ID,
                Communications: newnote,
                CommunicationInitiatorId: Currentusers.length > 0 ? Currentusers[0].ID : null
            };
            this.props.sharepointservice.addTicketNotes(tnote).then((tdata) => {
                if (user) {
                    const ticketnote: any = {
                        author: user.Title,
                        avatar: user.pictureurl,
                        content: newnote,
                        datetime: new Date().toString()
                    };
                    notesdata.push(ticketnote);
                    //notesdata.sort((a:any,b:any)=> new Date(a.datetime) - new Date(b.datetime));
                    this.setState({ notesdata: notesdata });
                }
            });
        }
    }

    public internalnoteChange = (e) => {
        this.setState({ newnote: e.currentTarget.value });
    }

    public closingcommentsChange = (e) => {
        this.setState({ closingComments: e.currentTarget.value });
    }

    public tstatuschange = (value) => {
        if (value == "7" || value == "12" || value == "14") {
            this.setState({ isStatusClosed: true, newstatus: value });
        } else {
            this.setState({ isStatusClosed: false, newstatus: value });
        }
    }

    public ateamchange = (value) => {
        this.setState({ newat: value });
    }

    public descriptionChange = (e) => {
        this.setState({ ticketdescription: e });
        return e;
    }

    public titleChange = (e) => {
        this.setState({ tickettitle: e.currentTarget.value });
    }

    public urgencyChange = (value) => {
        this.setState({ ticketurgency: value });
    }

    public impactChange = (value) => {
        this.setState({ ticketimpact: value });
    }

    public onSubmit = () => {
        debugger;
        this.setState({ loading: true });
        let nin:string=null;
        const { assignedperson, newat, newstatus, closingComments, tickettitle, ticketdescription, ticketimpact, ticketurgency,requester } = this.state;
        let assignedpersonid: IUserDetails[] = [];
        if (typeof assignedperson != "undefined" && assignedperson.length > 0) {
            assignedpersonid = this.props.sharepointservice._lusers.filter(i => i.Email == assignedperson[0].secondaryText);
        } else {
            assignedpersonid = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.selectedTicket.AssignedPerson);
        }

        let requesterid: IUserDetails[] = [];
        if (typeof requester != "undefined" && requester.length > 0) {
            requesterid = this.props.sharepointservice._lusers.filter(i => i.Email == requester[0].secondaryText);
        } else {
            requesterid = this.props.sharepointservice._lusers.filter(i => i.Title == this.props.selectedTicket.Requester);
        }

        const {servicegroupid,serviceid,subcategoryid,relatedCiid}=this.state.classificationvalues;
        let rci:any=[];
        if(relatedCiid.length>0){
            rci=[relatedCiid];
        }
        console.log(this.state.internalnotesvalues);
        if(this.state.internalnotesvalues!=undefined){
        const {newinternalnote,isAssignedPerson,isAssignedTeam,isAssignedStaff,staffusers}=this.state.internalnotesvalues;
            nin=newinternalnote;
            if(isAssignedPerson||isAssignedTeam||isAssignedStaff){
                let usersnotify:any[]=[];
                if(isAssignedStaff){
                    staffusers.forEach(user => {
                        const suser=this.props.sharepointservice._lusers.filter(i=>i.Email==user.secondaryText);
                        if(suser.length>0){usersnotify.push(suser[0].ID);}
                    });
                }
    
                const newnote={
                    "__metadata":{"type":"SP.Data.TicketNotesListItem"},
                    Title:tickettitle,
                    Note:newinternalnote,
                    RelatedTicketId:this.props.selectedTicket.ID,
                    TeamToNotifyId:isAssignedTeam?newat:null,
                    NoteAuthorId:isAssignedPerson?(assignedpersonid.length > 0 ? assignedpersonid[0].ID : null):null,
                    UsersToNotifyId:{
                        results:usersnotify
                    }
                };
    
                this.props.sharepointservice.AddTicketInternalNotes(newnote).then((ndata)=>{
                    console.log(ndata);
                });
            }
    }
        const updateticket = {
            "__metadata":{"type":"SP.Data.TicketsListItem"},
            AssignedPersonId: assignedpersonid.length > 0 ? assignedpersonid[0].ID : null,
            TicketsStatusId: newstatus,
            AssignedTeamId: newat,
            CloseRemark: closingComments,
            Impact: ticketimpact,
            Urgency: ticketurgency,
            Description:ticketdescription,
            Title: tickettitle,
            RequesterId:requesterid.length>0?requesterid[0].ID:null,
            ServiceGroupsId:servicegroupid,
            RelatedServicesId:serviceid,
            RelatedCategoriesId:subcategoryid,
            Notes:nin,
            RelatedCIsId:{'results':rci}
        };
        console.log("update obj",updateticket);
        this.props.sharepointservice.updateTicketDetails(updateticket, this.props.selectedTicket.ID).then((data) => {
            this.props.refreshticketsdata(undefined);
            this.setState({ isdrawervisible: false, loading: false });
        }).catch((ex) => {
            console.log("From ticket update componenet: Error while updating ticket details; ", ex);
            this.setState({ errorMessage: "error while updating ticket details", iserror: true, loading: false });
        });
        
    }

    public getclassificationvalues=(cvalues)=>{
        this.setState({classificationvalues:cvalues});
    }

    public getinternalnotevalues=(invalues)=>{
        this.setState({internalnotesvalues:invalues});
    }

    public render(): React.ReactElement<IItsm360EditTicketProps> {
        const { Priority, Prioritycolor, Title, Status, AssignedTeam, AssignedPerson, Requester } = this.props.selectedTicket;
        const {orderdetails}=this.state;
        //Setting the Selected status
        const ss = this.props.status.filter(i => i.Title == Status);
        const statusid = ss.length > 0 ? ss[0].ID : "Select a status";
        //Setting the selected Assigned Team if any
        const at = this.props.teams.filter(i => i.Title == AssignedTeam);
        const ateamid = at.length > 0 ? at[0].ID : "Select a Team";
        //Setting the selected Assigned Person If any
        const ap = typeof AssignedPerson != "undefined" ? this.props.sharepointservice._lusers.filter(i => i.Email == AssignedPerson) : null;
        const selectedap = ap && ap.length > 0 ? ap[0].Email : "";
        const rr = this.props.sharepointservice._lusers.filter(i => i.Title == Requester);
        const selectedrr = rr && rr.length > 0 ? rr[0].Email : "";

        return (
            <div style={{ display: "inline" }}>
                <Button type="link" onClick={this.editTicketClick}>{this.props.tictitle}</Button>
                <Drawer
                    title={`${this.props.selectedTicket.Title}, ID ${this.props.selectedTicket.ID}`}
                    width="80%"
                    onClose={this.handleClose}
                    visible={this.state.isdrawervisible}
                    destroyOnClose={true}
                >
                    <div>
                        <Row align="middle" type="flex" justify="space-between">
                            <Col span={5}>
                                <Form.Item label="Assigned Team" >
                                    <Select placeholder="Select a Team" defaultValue={ateamid} onChange={this.ateamchange} style={{ width: "90%" }}>
                                        {this.props.teams.map((team: ITeam, index) => <Option value={team.ID} key={index}>{team.Title}</Option>)}
                                    </Select>
                                </Form.Item>
                            </Col>
                            <Col span={7}>
                                <Form.Item label="Assigned Person" style={{ width: "90%" }}>
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
                                        defaultSelectedUsers={[selectedap]}
                                        resolveDelay={200} />
                                </Form.Item>
                            </Col>
                            <Col span={5}>
                                <Form.Item label="Status">
                                    <Select placeholder="Select Status" defaultValue={statusid} onChange={this.tstatuschange} style={{ width: "90%" }} >
                                        {this.props.status.map((stat: Istatus, index) => <Option value={stat.ID} key={index}>{stat.Title}</Option>)}
                                    </Select>
                                </Form.Item>
                            </Col>
                            <Col span={3}>
                                <Form.Item label="Impact">
                                    <Select placeholder="Select Impact" onChange={this.impactChange} value={this.state.ticketimpact}>
                                        <Option value="Low">Low</Option>
                                        <Option value="Mid">Mid</Option>
                                        <Option value="High">High</Option>
                                    </Select>
                                </Form.Item>
                            </Col>
                            <Col span={3}>
                                <Form.Item label="Urgency">
                                    <Select placeholder="Select Urgency" onChange={this.urgencyChange} value={this.state.ticketurgency}>
                                        <Option value="Low">Low</Option>
                                        <Option value="Mid">Mid</Option>
                                        <Option value="High">High</Option>
                                    </Select>
                                </Form.Item>
                            </Col>
                            <Col span={1}>
                                <Form.Item label="Priority">
                                    <Avatar shape="square" style={{background: Prioritycolor,color:"#fff"}}>{Priority}</Avatar>
                                </Form.Item>
                            </Col>
                        </Row>
                        <Row align="middle" type="flex" justify="space-between">
                            <Col span={12} style={{borderRight: "1px solid #e8e8e8"}}>
                            <Divider orientation="left">Update Ticket Information</Divider>
                            </Col>
                            <Col span={12}>
                            <Divider orientation="left">Summary</Divider>
                            </Col>
                        </Row>
                        <Row align="top" type="flex" justify="space-between">
                            <Col span={12} style={{borderRight: "1px solid #e8e8e8"}}>
                                {/* <Divider orientation="left">Update Ticket Information</Divider> */}
                                <div className="card-container">
                                    <Tabs type="card">
                                        <TabPane tab="Details" key="1">
                                            <Form.Item label="Title">
                                                <Input style={{ width: "40%" }} onChange={this.titleChange} value={this.state.tickettitle} />
                                            </Form.Item>
                                            <Form.Item label="Requestor" style={{ width: "50%" }}>
                                                <PeoplePicker
                                                    context={this.props.ppcontext}
                                                    titleText=""
                                                    personSelectionLimit={1}
                                                    showtooltip={true}
                                                    isRequired={false}
                                                    disabled={false}
                                                    selectedItems={this._getrequesterpeopleItems}
                                                    showHiddenInUI={false}
                                                    principalTypes={[PrincipalType.User]}
                                                    defaultSelectedUsers={[selectedrr]}
                                                    resolveDelay={200} />
                                            </Form.Item>
                                            <Form.Item label="Description">
                                                {/* <TextArea rows={3} style={{ width: "60%" }} onChange={this.descriptionChange} value={this.state.ticketdescription} /> */}
                                                <div style={{ border: "1px solid #d9d9d9",maxHeight:"150px",overflow:"auto",width:"80%"}} key={this._count}>
                                                    <RichText isEditMode={true} value={this.state.ticketdescription} onChange={this.descriptionChange} />   </div>
                                            </Form.Item>
                                            <Form.Item label="Attachments">
                                                <List
                                                    size="small"
                                                    bordered
                                                    style={{ width: "40%", height: "auto" }}
                                                    dataSource={this.state.ticketattachments}
                                                    renderItem={item => <List.Item><a href={item.attachurl} target="_blank">{item.filename}</a></List.Item>}
                                                />
                                            </Form.Item>
                                        </TabPane>
                                        <TabPane tab="Classification" key="2">
                                            {typeof this.state.classificationvalues!="undefined"?<Itsm360Classification sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket} defaultvalues={this.state.classificationvalues} getclassificationvalues={this.getclassificationvalues} />:""}

                                        </TabPane>
                                        <TabPane tab="Notes" key="3">
                                            <Itsm360InternalNotes sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket} getinternalnotevalues={this.getinternalnotevalues} ppcontext={this.props.ppcontext} />
                                        </TabPane>
                                        <TabPane tab="Conversation" key="4">
                                                <Itsm360Conversation spservice={this.props.sharepointservice} classificationvals={this.state.classificationvalues} selectedTicket={this.props.selectedTicket} />
                                        </TabPane>
                                        <TabPane tab="Emails" key="5">
                                            <Itsm360Emails sharepointservice={this.props.sharepointservice} selectedTicket={this.props.selectedTicket}></Itsm360Emails>
                                        </TabPane>
                                        <TabPane tab="Resolve" key="6">
                                            <Form layout="vertical">
                                                <Form.Item label="Closing Remarks">
                                                    <TextArea placeholder="Make a note" rows={3} style={{ width: "60%" }} disabled={!this.state.isStatusClosed} onChange={this.closingcommentsChange} value={this.state.closingComments} />
                                                </Form.Item>
                                            </Form>

                                            <Itsm360SubTasks sharepointservice={this.props.sharepointservice} ticketid={this.props.selectedTicket.ID} ppcontext={this.props.ppcontext} teams={this.props.teams} />        
                                        </TabPane>
                                    </Tabs>
                                </div>

                            </Col>
                            <Col span={12}>
                              <div dangerouslySetInnerHTML={{__html: this.state.notificationsummary}} />
                             <Divider orientation="left">Order Details</Divider>
                               {orderdetails!=null && orderdetails.length>0?<Itsm360OrderDetails spservice={this.props.sharepointservice} orderdetails={orderdetails} />:""}             
                            </Col>
                        </Row>
                    </div>
                    <div className="itsmdrawerbuttons">
                        <Button onClick={this.handleClose} style={{ marginRight: 8 }}>
                            Cancel
                        </Button>
                        <Button onClick={this.onSubmit} type="primary" loading={this.state.loading}>
                            Submit
                        </Button>
                    </div>
                </Drawer>
            </div>
        );
    }
}
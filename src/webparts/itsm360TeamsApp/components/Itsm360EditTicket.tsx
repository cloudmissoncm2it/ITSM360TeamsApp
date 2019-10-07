import * as React from 'react';
import { Drawer, Button, Row, Col, Divider, Tabs, Form, Select, Input, List,Comment } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import * as moment from 'moment';
import { ITeam } from '../model/ITeam';
import { Istatus } from '../model/Istatus';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserDetails } from '../model/IUserDetails';
const { TabPane } = Tabs;
const { Option } = Select;
const { TextArea } = Input;

export interface IItsm360EditTicketProps {
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
    teams?:ITeam[];
    status?:Istatus[];
    ppcontext:WebPartContext;
}

export interface IItsm360EditTicketState {
    isdrawervisible?: boolean;
    modalsave?: boolean;
    iserror?:boolean;
    errorMessage?: string;
    assignedperson?:any[];
    notesdata?: any[];
    newnote?:string;
    isStatusClosed?:boolean;
    newat?:string;
    newstatus?:string;
    closingComments?:string;
    loading?:boolean;
    ticketattachments?:any[];
}

export class Itsm360EditTicket extends React.Component<IItsm360EditTicketProps, IItsm360EditTicketState>{

    constructor(props: IItsm360EditTicketProps) {
        super(props);
        this.state = {
            isdrawervisible: false,
            modalsave: false,
            iserror: false,
            notesdata:[],
            isStatusClosed:false,
            loading:false,
            ticketattachments:[]
        };
    }

    public handleClose = (e) => {
        this.setState({ isdrawervisible: false });
    }

    public _getPeoplePickerItems=(people: any[])=> {
        this.setState({assignedperson:people});
    }

    public editTicketClick = (e) => {
        this.setState({ isdrawervisible: true });
        const {Status,ID}=this.props.selectedTicket;
        const spservice=this.props.sharepointservice;
        debugger;
        if(Status.indexOf("Closed")>-1){
            this.setState({isStatusClosed:true});
        }
        console.log(this.props.selectedTicket);
        spservice.getTicketNotes(ID).then((notesdata) => {
            this.setState({ notesdata: notesdata });
        });
        spservice.getTicketAttachment(ID).then((ticketattach) => {
            this.setState({ ticketattachments: ticketattach });
        });
        
    }

    public postinternalnotes=()=>{
        const { newnote,notesdata } = this.state;
        debugger;
        if (typeof newnote != "undefined") {
            const Currentusers: IUserDetails[] = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.sharepointservice._currentuser.email);
            const user:IUserDetails=Currentusers.length > 0 ? Currentusers[0] : null;
            const tnote = {
                TicketIDId: this.props.selectedTicket.ID,
                Communications: newnote,
                CommunicationInitiatorId: Currentusers.length > 0 ? Currentusers[0].ID : null
            };
            this.props.sharepointservice.addTicketNotes(tnote).then((tdata) => {
                debugger;
                if(user){
                    const ticketnote:any={
                        author:user.Title,
                        avatar:user.pictureurl,
                        content:newnote,
                        datetime:new Date().toString()
                    };
                    notesdata.push(ticketnote);
                    //notesdata.sort((a:any,b:any)=> new Date(a.datetime) - new Date(b.datetime));
                    this.setState({notesdata:notesdata});
                }
            });
        }
    }

    public internalnoteChange = (e) => {
        this.setState({ newnote: e.currentTarget.value });
    }

    public closingcommentsChange=(e)=>{
        this.setState({ closingComments: e.currentTarget.value });
    }

    public tstatuschange=(value)=>{
        if(value=="7"||value=="12"||value=="14"){
            this.setState({isStatusClosed:true,newstatus:value});
        }else{
            this.setState({isStatusClosed:false,newstatus:value});
        }
    }

    public ateamchange=(value)=>{
        this.setState({newat:value});
    }

    public onSubmit=()=>{
        this.setState({loading:true});
        const {assignedperson,newat,newstatus,closingComments}=this.state;
        let assignedpersonid:IUserDetails[]=[];
        if(typeof assignedperson !="undefined" && assignedperson.length>0){
        assignedpersonid=this.props.sharepointservice._lusers.filter(i=>i.Email==assignedperson[0].secondaryText);
        }else{
            assignedpersonid=this.props.sharepointservice._lusers.filter(i=>i.Email==this.props.selectedTicket.AssignedPerson);
        }
        const updateticket={
            AssignedPersonId:assignedpersonid.length>0?assignedpersonid[0].ID:null,
            TicketsStatusId:newstatus,
            AssignedTeamId:newat,
            CloseRemark:closingComments
        };

        this.props.sharepointservice.updateTicketDetails(updateticket,this.props.selectedTicket.ID).then((data)=>{
            this.setState({isdrawervisible:false,loading:false});
        }).catch((ex) => {
            console.log("From ticket update componenet: Error while updating ticket details; ", ex);
            this.setState({errorMessage:"error while updating ticket details",iserror:true,loading:false});
        });
    }

    public render(): React.ReactElement<IItsm360EditTicketProps> {
        const {lastmodified,lastmodifiedby,Priority,Prioritycolor,Title,Status,AssignedTeam,AssignedPerson,Requester}=this.props.selectedTicket;
        //Setting the Selected status
        const ss=this.props.status.filter(i=>i.Title==Status);
        const statusid=ss.length>0?ss[0].ID:"Select a status";
        //Setting the selected Assigned Team if any
        const at=this.props.teams.filter(i=>i.Title==AssignedTeam);
        const ateamid=at.length>0?at[0].ID:"Select a Team";
        //Setting the selected Assigned Person If any
        const ap=typeof AssignedPerson !="undefined"?this.props.sharepointservice._lusers.filter(i=>i.Email==AssignedPerson):null;
        const selectedap=ap && ap.length>0?ap[0].Email:"";
        return (
            <div>
                <Button type="primary" icon="edit" size="small" onClick={this.editTicketClick} />
                <Drawer
                    title={`Ticket Information, ID ${this.props.selectedTicket.ID}`}
                    width="60%"
                    onClose={this.handleClose}
                    visible={this.state.isdrawervisible}
                    destroyOnClose={true}
                >
                    <div>
                        <Row align="middle" type="flex" justify="space-between">
                            <Col span={9}>
                                <h4>{Title}</h4>
                            </Col>
                            <Col span={15}>
                                <Row>
                                    <Col span={12}>
                                        <div className="itsmdrawer">
                                            <p className="itsmdrawertitle">
                                                Current Status:
                                        </p>
                                            {Status}
                                        </div>
                                    </Col>
                                    <Col span={12}>
                                        <div className="itsmdrawer">
                                            <p className="itsmdrawertitle">
                                                Priority:
                                            </p>
                                            <p style={{ background: Prioritycolor, textAlign: 'center', color: '#ffffff', padding: '3%', display: 'inline' }}>
                                                {Priority}
                                            </p>
                                        </div>
                                    </Col>
                                </Row>
                                <Row>
                                    <Col span={12}>
                                        <div className="itsmdrawer">
                                            <p className="itsmdrawertitle">
                                                Last Modified:
                                        </p>
                                            {`${moment(lastmodified).fromNow()} by ${lastmodifiedby}`}
                                        </div>
                                    </Col>
                                    <Col span={12}>
                                        <div className="itsmdrawer">
                                            <p className="itsmdrawertitle">
                                                Time to Resolution:
                                        </p>
                                            -26 d 7h 9m Breached
                                        </div>
                                    </Col>
                                </Row>
                            </Col>
                        </Row>
                        <Row align="middle" type="flex" justify="space-between">
                            
                            <Divider orientation="left">Update Ticket Information</Divider>
                            <div className="card-container" style={{width:"75%"}}>
                                <Tabs type="card">
                                    <TabPane tab="Details" key="1">
                                    <Form.Item label="Title">
                                        <Input defaultValue={Title} disabled={true} style={{ width: "40%" }} />
                                    </Form.Item>
                                    <Form.Item label="Requestor">
                                        <Input defaultValue={Requester} disabled={true} style={{ width: "40%" }} />
                                    </Form.Item>
                                    <Form.Item label="Description">
                                        <TextArea defaultValue={Title} rows={3} style={{ width: "60%" }} disabled={true} />
                                    </Form.Item>
                                    <Form.Item label="Attachments">
                                    <List
                                        size="small"
                                        bordered
                                        style={{width:"40%",height:"auto"}}
                                        dataSource={this.state.ticketattachments}
                                        renderItem={item => <List.Item><a href={item.attachurl} target="_blank">{item.filename}</a></List.Item>}
                                        />
                                    </Form.Item>
                                    </TabPane>
                                    <TabPane tab="Assign" key="2">
                                        <Form layout="vertical">
                                            <Form.Item label="Assigned Team">
                                                <Select placeholder="Select a Team" style={{ width: "40%" }}defaultValue={ateamid} onChange={this.ateamchange}>
                                                {this.props.teams.map((team: ITeam, index) => <Option value={team.ID} key={index}>{team.Title}</Option>)}
                                                </Select>
                                            </Form.Item>
                                            <Form.Item label="Assigned Person" style={{width:"50%"}}>
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
                                        </Form>
                                    </TabPane>
                                    <TabPane tab="Notes" key="3">
                                    <Form.Item>
                                        <TextArea placeholder="Post a message" rows={3} style={{ width: "60%" }} onChange={this.internalnoteChange} value={this.state.newnote} />
                                        <div style={{marginLeft:"50%"}}>
                                            <Button type="primary" icon="message" size="small" onClick={this.postinternalnotes}>Post</Button>
                                        </div>
                                    </Form.Item>
                                    <List
                                        className="comment-list"
                                        header={`${this.state.notesdata.length} replies`}
                                        itemLayout="horizontal"
                                        dataSource={this.state.notesdata}
                                        renderItem={item => (
                                            <li>
                                                <Comment
                                                    author={item.author}
                                                    avatar={item.avatar}
                                                    content={item.content}
                                                    datetime={moment(item.datetime).fromNow()}
                                                />
                                            </li>
                                        )}
                                    />
                                    </TabPane>
                                    <TabPane tab="Resolve" key="4">
                                        <Form layout="vertical">
                                        <Form.Item label="Status">
                                                <Select placeholder="Select Status" style={{ width: "40%" }} defaultValue={statusid} onChange={this.tstatuschange} >
                                                {this.props.status.map((stat: Istatus, index) => <Option value={stat.ID} key={index}>{stat.Title}</Option>)}
                                                </Select>
                                        </Form.Item>
                                        <Form.Item label="Closing Remarks">
                                                <TextArea placeholder="Make a note" rows={3} style={{ width: "60%" }} disabled={!this.state.isStatusClosed} onChange={this.closingcommentsChange} value={this.state.closingComments} />
                                        </Form.Item>
                                        </Form>
                                    </TabPane>
                                </Tabs>
                            </div>
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
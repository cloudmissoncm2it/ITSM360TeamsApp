import * as React from 'react';
import { Drawer, Button, Row, Col, Divider, Tabs, Form, Select, Input, List, Comment, Cascader,Descriptions } from 'antd';
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
    teams?: ITeam[];
    status?: Istatus[];
    ppcontext: WebPartContext;
    tictitle: string;
}

export interface IItsm360EditTicketState {
    isdrawervisible?: boolean;
    modalsave?: boolean;
    iserror?: boolean;
    errorMessage?: string;
    assignedperson?: any[];
    requester?: any[];
    notesdata?: any[];
    internalnotes?: any[];
    newinternalnote?: string;
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
    cascaderoptions?: any[];
    //cascaderdefault?:any[];
    servicegroup?: string;
    service?: string;
    subcategory?: string;
}

export class Itsm360EditTicket extends React.Component<IItsm360EditTicketProps, IItsm360EditTicketState>{

    private _sgtitle:string;
    private _setitle:string;
    private _scategorytitle:String;

    constructor(props: IItsm360EditTicketProps) {
        super(props);
        this.state = {
            isdrawervisible: false,
            modalsave: false,
            iserror: false,
            notesdata: [],
            internalnotes: [],
            isStatusClosed: false,
            loading: false,
            ticketattachments: [],
            tickettitle: this.props.selectedTicket.Title,
            ticketimpact: "Low",
            ticketurgency: "Low"
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
        //console.log(this.props.selectedTicket);
        spservice.getTicketNotes(ID).then((notesdata) => {
            this.setState({ notesdata: notesdata });
        });
        spservice.getTicketAttachment(ID).then((ticketattach) => {
            this.setState({ ticketattachments: ticketattach });
        });

        spservice.getTicketDetails(ID).then((ticketdata) => {
            this.setState({
                ticketdescription: ticketdata.Description,
                ticketurgency: ticketdata.Urgency,
                ticketimpact: ticketdata.Impact
            });
            this._sgtitle=ticketdata.ServiceGroups.Title;
            this._setitle=ticketdata.RelatedServices.Title;
            this._scategorytitle=ticketdata.RelatedCategories.Title;
            //const dv=[ticketdata.ServiceGroups.ID,ticketdata.RelatedServices.ID,ticketdata.RelatedCategories.ID];
            spservice.getlookupdatanew().then((cddata) => {
                this.setState({
                    cascaderoptions: cddata,
                    servicegroup: ticketdata.ServiceGroups.ID,
                    service: ticketdata.RelatedServices.ID,
                    subcategory:ticketdata.RelatedCategories.ID
                });
            });
        });

        spservice.getTicketInternalNotes(ID).then((notesdata) => {
            this.setState({ internalnotes: notesdata });
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

    public ticketnoteChange = (e) => {
        this.setState({ newinternalnote : e.currentTarget.value });
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
        this.setState({ ticketdescription: e.currentTarget.value });
    }

    public titleChange = (e) => {
        this.setState({ tickettitle: e.currentTarget.value });
    }

    public cascaderChange = (e) => {
        console.log(e);
        if (e.length == 3) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: e[2]
            });
        } else if (e.length == 2) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: "-1"
            });
        } else if (e.length == 1) {
            this.setState({
                servicegroup: e[0],
                service: "-1",
                subcategory: "-1"
            });
        } else {
            this.setState({
                servicegroup: "-1",
                service: "-1",
                subcategory: "-1"
            });
        }
    }

    public urgencyChange = (value) => {
        this.setState({ ticketurgency: value });
    }

    public impactChange = (value) => {
        this.setState({ ticketimpact: value });
    }

    public onSubmit = () => {
        this.setState({ loading: true });
        const { assignedperson, newat, newstatus, closingComments, tickettitle, ticketdescription, ticketimpact, ticketurgency,requester,servicegroup,service,subcategory,newinternalnote } = this.state;
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

        const updateticket = {
            AssignedPersonId: assignedpersonid.length > 0 ? assignedpersonid[0].ID : null,
            TicketsStatusId: newstatus,
            AssignedTeamId: newat,
            CloseRemark: closingComments,
            Impact: ticketimpact,
            Urgency: ticketurgency,
            Description:ticketdescription,
            Title: tickettitle,
            RequesterId:requesterid.length>0?requesterid[0].ID:null,
            ServiceGroupsId:servicegroup,
            RelatedServicesId:service,
            RelatedCategoriesId:subcategory,
            Notes:newinternalnote
        };

        this.props.sharepointservice.updateTicketDetails(updateticket, this.props.selectedTicket.ID).then((data) => {
            this.setState({ isdrawervisible: false, loading: false });
        }).catch((ex) => {
            console.log("From ticket update componenet: Error while updating ticket details; ", ex);
            this.setState({ errorMessage: "error while updating ticket details", iserror: true, loading: false });
        });
    }

    public render(): React.ReactElement<IItsm360EditTicketProps> {
        const { Priority, Prioritycolor, Title, Status, AssignedTeam, AssignedPerson, Requester } = this.props.selectedTicket;
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
                    width="75%"
                    onClose={this.handleClose}
                    visible={this.state.isdrawervisible}
                    destroyOnClose={true}
                >
                    <div>
                        <Row align="middle" type="flex" justify="space-between">
                            <Col span={5}>
                                <Form.Item label="Assigned Team" style={{ width: "50%" }}>
                                    <Select placeholder="Select a Team" defaultValue={ateamid} onChange={this.ateamchange}>
                                        {this.props.teams.map((team: ITeam, index) => <Option value={team.ID} key={index}>{team.Title}</Option>)}
                                    </Select>
                                </Form.Item>
                            </Col>
                            <Col span={7}>
                                <Form.Item label="Assigned Person" style={{ width: "80%" }}>
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
                                    <Select placeholder="Select Status" defaultValue={statusid} onChange={this.tstatuschange} >
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
                                    <p style={{ background: Prioritycolor, textAlign: 'center', color: '#ffffff', padding: '3%', display: 'inline' }}>
                                        {Priority}
                                    </p>
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
                                                <TextArea rows={3} style={{ width: "60%" }} onChange={this.descriptionChange} value={this.state.ticketdescription} />
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
                                            <Form layout="vertical">
                                                <Form.Item label="Classification" style={{ width: "50%" }}>
                                                    <Cascader options={this.state.cascaderoptions} onChange={this.cascaderChange} defaultValue={[this.state.servicegroup,this.state.service,this.state.subcategory]} />
                                                </Form.Item>
                                            </Form>
                                        </TabPane>
                                        <TabPane tab="Notes" key="3">
                                            <Form.Item>
                                                <TextArea placeholder="Post a Note" rows={3} style={{ width: "60%" }} onChange={this.ticketnoteChange} value={this.state.newinternalnote} />
                                            </Form.Item>
                                            <List
                                                className="comment-list"
                                                header={`${this.state.internalnotes.length} note`}
                                                itemLayout="horizontal"
                                                dataSource={this.state.internalnotes}
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
                                        <TabPane tab="Conversation" key="4">
                                            <Form.Item>
                                                <TextArea placeholder="Post a message" rows={3} style={{ width: "60%" }} onChange={this.internalnoteChange} value={this.state.newnote} />
                                                <div style={{ marginLeft: "50%" }}>
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
                                        <TabPane tab="Resolve" key="5">
                                            <Form layout="vertical">
                                                <Form.Item label="Closing Remarks">
                                                    <TextArea placeholder="Make a note" rows={3} style={{ width: "60%" }} disabled={!this.state.isStatusClosed} onChange={this.closingcommentsChange} value={this.state.closingComments} />
                                                </Form.Item>
                                            </Form>
                                        </TabPane>
                                    </Tabs>
                                </div>

                            </Col>
                            <Col span={12}>
                             <Descriptions bordered style={{marginLeft:"1%"}} size="small">
                                 <Descriptions.Item label="Subject" span={2}>{this.state.tickettitle}</Descriptions.Item>
                                 <Descriptions.Item label="Status">{ss[0].Title}</Descriptions.Item>
                                 <Descriptions.Item label="Requestor" span={3}>{Requester}</Descriptions.Item>
                                 <Descriptions.Item label="Description" span={3}>{this.state.ticketdescription}</Descriptions.Item>
                                 <Descriptions.Item label="Service Group" span={3}>{this._sgtitle}</Descriptions.Item>
                                 <Descriptions.Item label="Service">{this._setitle}</Descriptions.Item>
                                 <Descriptions.Item label="Category">{this._scategorytitle}</Descriptions.Item>
                             </Descriptions>
                             <Divider orientation="left">Order Details</Divider>
                                            
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
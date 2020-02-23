import * as React from 'react';
import { Button, List, Drawer, Input } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";

export interface IItsm360EmailsProps {
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
}

export interface IItsm360EmailsState {
    emails?: any[];
    showNewEmailDrawer?: boolean;
    requesterEmail?: string;
    senderEmail?: string;
    to?: string;
    cc?: string;
    subject?: string;
    body?: string;
    loading?: boolean;
}

export class Itsm360Emails extends React.Component<IItsm360EmailsProps, IItsm360EmailsState>{
    constructor(props: IItsm360EmailsProps) {
        super(props);
        this.state = {
            emails: [],
            showNewEmailDrawer: false,
            requesterEmail: "",
            senderEmail: "",
            to: "",
            cc: "",
            subject: "",
            body: "",
            loading: false
        };
    }

    public componentDidMount() {
        const spservice = this.props.sharepointservice;
        spservice.getTicketEmails(this.props.selectedTicket.ID).then((emailsData) => {
            this.setState({
                emails: emailsData
            });
        });
        spservice.getTicketRequesterAndSenderEmails(this.props.selectedTicket.ID).then((ticketDetails) => {
            this.setState({
                requesterEmail: ticketDetails.Requester,
                senderEmail: ticketDetails.Sender
            });
        });
    }

    public render(): React.ReactElement<IItsm360EmailsProps> {
        return (
            <div>
                <Button onClick={ this.newEmailClick }>New email</Button>
                <List
                    className="comment-list"
                    itemLayout="horizontal"
                    dataSource={this.state.emails}
                    renderItem={item => (
                        <li>
                            <div><b>To:</b> {item.Email}</div>
                            <div style={(item.Cc != "" && item.Cc != null) ? {display:"block"} : {display:"none"}}><b>Cc:</b> {item.Cc}</div>
                            <div><b>Subject:</b> {item.Title}</div>
                            <div><b>Date:</b> {item.Created}</div>
                            <div dangerouslySetInnerHTML={{ __html: item.Message }}></div>
                            <hr/>
                        </li>
                    )}
                />
                <Drawer
                    title="New email"
                    width="50%"
                    onClose={this.handleClose}
                    visible={this.state.showNewEmailDrawer}
                    destroyOnClose={true}
                >
                    <div>
                        <div>
                            <table style={{width:"100%"}}>
                                <tr>
                                    <td>
                                        <span className="newEmailLabel">To&nbsp;<span style={{color:"red"}}>*</span></span>
                                    </td>
                                    <td style={{textAlign:"right"}}>
                                        <Button icon="plus" onClick={ this.addReqesterEmailTo } style={this.state.requesterEmail != "" ? {display:"inline-block",paddingRight:"5px"} : {display:"none"}}>Add requester email</Button>&nbsp;
                                        <Button icon="plus" onClick={ this.addSenderEmailTo } style={this.state.senderEmail != "" ? {display:"inline-block"} : {display:"none"}}>Add sender email</Button>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <div style={{paddingTop:"5px",paddingBottom:"5px"}}>
                            <Input style={{ width: "80%" }} value={this.state.to} onChange={this.toChange} />
                        </div>
                        <div>
                            <table style={{width:"100%"}}>
                                <tr>
                                    <td>
                                        <span className="newEmailLabel">Cc</span>
                                    </td>
                                    <td style={{textAlign:"right"}}>
                                        <Button icon="plus" onClick={ this.addReqesterEmailCc } style={this.state.requesterEmail != "" ? {display:"inline-block",paddingRight:"5px"} : {display:"none"}}>Add requester email</Button>&nbsp;
                                        <Button icon="plus" onClick={ this.addSenderEmailCc } style={this.state.senderEmail != "" ? {display:"inline-block"} : {display:"none"}}>Add sender email</Button>
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <div style={{paddingTop:"5px",paddingBottom:"5px"}}>
                            <Input style={{ width: "80%" }} value={this.state.cc} onChange={this.ccChange} />
                        </div>
                        <div>
                            <span className="newEmailLabel">Subject</span>
                        </div>
                        <div style={{paddingTop:"5px",paddingBottom:"5px"}}>
                            <Input style={{ width: "80%" }} onChange={this.subjectChange} />
                        </div>
                        <div>
                        <span className="newEmailLabel">Message</span>
                        </div>
                        <div style={{marginTop:"5px",border:"1px solid rgb(217,217,217)",borderRadius:"4px"}}>
                            <RichText value="" onChange={(txt) => { this.onBodyChange(txt); return txt; }}/>
                        </div>
                        <div className="itsmdrawerbuttons">
                            <Button onClick={this.handleClose} style={{ marginRight: 8 }}>
                                Cancel
                            </Button>
                            <Button onClick={this.newEmailSaveClick} type="primary" loading={this.state.loading} disabled={this.state.to == ""}>
                                Send
                            </Button>
                        </div>
                    </div>
                </Drawer>
            </div>
        );
    }

    public newEmailClick = (e) => {
        this.setState({ showNewEmailDrawer: true });
    }

    public handleClose = (e) => {
        this.setState({ showNewEmailDrawer: false });
    }

    public toChange = (e) => {
        this.setState({ to: e.currentTarget.value });
    }

    public ccChange = (e) => {
        this.setState({ cc: e.currentTarget.value });
    }

    public subjectChange = (e) => {
        this.setState({ subject: e.currentTarget.value });
    }

    public onBodyChange = (newText: string) => {
        this.setState({ body: newText });
    }

    public addReqesterEmailTo = () => {
        debugger;
        if (this.state.requesterEmail != "" && this.state.to.indexOf(this.state.requesterEmail) == -1) {
            let toLocal = this.state.to;
            if (this.state.to != "") {
                toLocal += ";";
            }
            toLocal += this.state.requesterEmail;
            this.setState({ to: toLocal });
        }
    }

    private addSenderEmailTo = () => {
        if (this.state.senderEmail != "" && this.state.to.indexOf(this.state.senderEmail) == -1) {
            let toLocal = this.state.to;
            if (this.state.to != "") {
                toLocal += ";";
            }
            toLocal += this.state.senderEmail;
            this.setState({ to: toLocal });
        }
    }

    public addReqesterEmailCc = () => {
        debugger;
        if (this.state.requesterEmail != "" && this.state.cc.indexOf(this.state.requesterEmail) == -1) {
            let ccLocal = this.state.cc;
            if (this.state.cc != "") {
                ccLocal += ";";
            }
            ccLocal += this.state.requesterEmail;
            this.setState({ cc: ccLocal });
        }
    }

    private addSenderEmailCc = () => {
        if (this.state.senderEmail != "" && this.state.cc.indexOf(this.state.senderEmail) == -1) {
            let ccLocal = this.state.cc;
            if (this.state.cc != "") {
                ccLocal += ";";
            }
            ccLocal += this.state.senderEmail;
            this.setState({ cc: ccLocal });
        }
    }

    public newEmailSaveClick = (e) => {
        this.setState({ loading: true });
        const newEmailData = {
            Title: this.state.subject,
            Message: this.state.body,
            PlainTextMessage: this.state.body,
            Email: this.state.to,
            Cc: this.state.cc,
            Received: false,
            Read: true,
            RelatedItem: this.props.selectedTicket.ID,
            RelatedList: "Tickets",
            SendFullHtmlEmail: true
        };
        this.props.sharepointservice.addTicketEmail(newEmailData).then((tdata) => {
            const spservice = this.props.sharepointservice;
            spservice.getTicketEmails(this.props.selectedTicket.ID).then((emailsData) => {
                this.setState({
                    emails: emailsData,
                    showNewEmailDrawer: false,
                    loading: false,
                    to: "",
                    cc: ""
                });
            });
        });
    }
}
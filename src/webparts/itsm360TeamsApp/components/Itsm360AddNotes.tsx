import * as React from 'react';
import { Modal, Button, Icon, Form, Alert } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { ITicketItem } from '../model/ITicketItem';
import { IUserDetails } from '../model/IUserDetails';
import * as moment from 'moment';

export interface IItsm360AddNotesProps {
    visible: boolean;
    sharepointservice: sharepointservice;
    selectedTicket?: ITicketItem;
    selectedRowKeys?: string[];
}

export interface IItsm360AddNotesState {
    ismodalvisible: boolean;
    modalsave?: boolean;
    errorMessage?: boolean;
    newnote?: string;
}

export class Itsm360AddNotes extends React.Component<IItsm360AddNotesProps, IItsm360AddNotesState>{

    constructor(props: IItsm360AddNotesProps) {
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false,
            errorMessage: false
        };
    }

    public handleOk = (e) => {
        const { newnote } = this.state;
        this.setState({ modalsave: true });
        const Currentusers: IUserDetails[] = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.sharepointservice._currentuser.email);
        const curruser: any = Currentusers.length > 0 ? Currentusers[0].ID : null;
        if (typeof newnote != "undefined" && newnote.length > 0) {
            this.props.selectedRowKeys.forEach((ticketid) => {
                const tnote = {
                    TicketIDId: ticketid,
                    Communications: newnote,
                    CommunicationInitiatorId: curruser
                };
                this.props.sharepointservice.addTicketNotes(tnote).then((tdata) => {
                    console.log(`Notes updated for ticket ${ticketid} with the output ${tdata}`);
                });
            });
            this.setState({
                modalsave: false,
                ismodalvisible: false,
                errorMessage: false
            });
        } else {
            this.setState({
                errorMessage: true,
                modalsave: false
            });
        }
    }

    public handleCancel = (e) => {
        this.setState({ ismodalvisible: false });
    }

    public groupconversationclick = (e) => {
        this.setState({ ismodalvisible: true });
    }

    public descriptionChange = (value: string) => {
        this.setState({ newnote: value });
        return value;
    }

    public render(): React.ReactElement<IItsm360AddNotesProps> {

        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={this.groupconversationclick}>
                    <Icon type="user-add" />
                    Send Group Conversation
                </Button>
                <Modal title="Send Group Conversation"
                    visible={this.state.ismodalvisible}
                    onOk={this.handleOk}
                    onCancel={this.handleCancel}
                    okText="Update Notes"
                    confirmLoading={this.state.modalsave}
                >
                    <div>
                        {this.state.errorMessage ? <Alert type="error" style={{ marginBottom: "10px" }} closable message="No new notes added." /> : ""}

                        <Alert type="info" style={{ marginBottom: "10px" }}
                            message="Selected Ticket IDs"
                            description={this.props.selectedRowKeys.toString()}
                        />

                        <Form layout="horizontal" labelCol={{ span: 5 }} wrapperCol={{ span: 12 }}>
                            <Form.Item label="Notes">
                                <div style={{border:"1px solid #d9d9d9"}}>
                                <RichText value={this.state.newnote} onChange={this.descriptionChange} /></div>
                            </Form.Item>

                        </Form>
                    </div>
                </Modal>
            </div>
        );
    }
}
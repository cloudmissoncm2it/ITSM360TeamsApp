import * as React from 'react';
import { Modal, Button, Icon, List, Comment, Tooltip, Form, Input, Alert } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { IUserDetails } from '../model/IUserDetails';

export interface IItsm360AddNotesProps {
    visible: boolean;
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
}

export interface IItsm360AddNotesState {
    ismodalvisible: boolean;
    modalsave?: boolean;
    errorMessage?: boolean;
    notesdata?: any[];
    newnote?: string;
}

export class Itsm360AddNotes extends React.Component<IItsm360AddNotesProps, IItsm360AddNotesState>{

    constructor(props: IItsm360AddNotesProps) {
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false,
            errorMessage: false,
            notesdata: []
        };
    }

    public handleOk = (e) => {
        const { newnote } = this.state;
        this.setState({ modalsave: true });
        if (typeof newnote != "undefined") {
            const Currentusers: IUserDetails[] = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.sharepointservice._currentuser.email);
            const tnote = {
                TicketIDId: this.props.selectedTicket.ID,
                Communications: newnote,
                CommunicationInitiatorId: Currentusers.length > 0 ? Currentusers[0].ID : null
            };
            this.props.sharepointservice.addTicketNotes(tnote).then((tdata) => {
                this.setState({
                    modalsave: false,
                    ismodalvisible: false,
                    errorMessage: false
                });
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
        this.props.sharepointservice.getTicketNotes(this.props.selectedTicket.ID).then((notesdata) => {
            this.setState({ notesdata: notesdata });
        });
    }

    public descriptionChange = (e) => {
        this.setState({ newnote: e.currentTarget.value });
    }

    public render(): React.ReactElement<IItsm360AddNotesProps> {
        const { TextArea } = Input;

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
                        <Form.Item>
                            <TextArea rows={4} onChange={this.descriptionChange} value={this.state.newnote} />
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
                                        datetime={item.datetime}
                                    />
                                </li>
                            )}
                        />
                    </div>
                </Modal>
            </div>
        );
    }
}
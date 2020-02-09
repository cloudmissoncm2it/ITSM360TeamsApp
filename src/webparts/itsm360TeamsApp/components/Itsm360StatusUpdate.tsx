import * as React from 'react';
import { Modal, Alert, Button, Icon, Form } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { Istatus } from '../model/Istatus';

export interface IItsm360StatusUpdateProps {
    visible: boolean;
    sharepointservice: sharepointservice;
    selectedTicket?: ITicketItem;
    selectedRowKeys?: string[];
    status: Istatus[];
}

export interface IItsm360StatusUpdateState {
    ismodalvisible: boolean;
    modalsave?: boolean;
    statusid?: string;
    errorMessage?: boolean;
}

export class Itsm360StatusUpdate extends React.Component<IItsm360StatusUpdateProps, IItsm360StatusUpdateState>{

    constructor(props: IItsm360StatusUpdateProps) {
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false
        };
    }

    public handleOk = (e) => {
        const { statusid } = this.state;
        this.setState({ modalsave: true });
        if (typeof statusid != "undefined" && statusid != "") {
            this.props.selectedRowKeys.forEach((ticketid) => {
                this.props.sharepointservice.updateTicketStatus(ticketid, this.state.statusid).then((resp) => {
                    console.log(`update status for ticketid ${ticketid} is ${resp}`);
                });
            });
            this.setState({
                modalsave: false,
                ismodalvisible: false
            });
        } else {
            this.setState({
                modalsave: false,
                errorMessage: true
            });
        }
    }

    public handleCancel = (e) => {
        this.setState({ ismodalvisible: false });
    }

    public statusChange = (e) => {
        this.setState({ statusid: e.target.value });
    }

    public render(): React.ReactElement<IItsm360StatusUpdateProps> {
        const stats: Istatus[] = this.props.status;
        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={() => this.setState({ ismodalvisible: true })}>
                    <Icon type="reload" />
                    Update Status
                </Button>
                <Modal title="Update Status"
                    visible={this.state.ismodalvisible}
                    onOk={this.handleOk}
                    onCancel={this.handleCancel}
                    okText="Update Status"
                    confirmLoading={this.state.modalsave}
                >
                    <div>
                        {this.state.errorMessage ? <Alert type="error" style={{ marginBottom: "10px" }} closable message="No Status selected. Select a status before updating" /> : ""}
                        <Alert type="info" style={{ marginBottom: "10px" }}
                            message="Selected Ticket IDs"
                            description={this.props.selectedRowKeys.toString()}
                        />
                        <Form layout="horizontal" labelCol={{ span: 5 }} wrapperCol={{ span: 12 }}>
                            <Form.Item label="Status">
                                <select onChange={this.statusChange} className="ant-select-selection ant-select-selection--single">
                                    <option value="">-Select a Status</option>
                                    {stats.map((stat: Istatus, index) => {
                                        return <option value={stat.ID} key={index}>{stat.Title}</option>;
                                    })}
                                </select>
                            </Form.Item>
                        </Form>
                    </div>
                </Modal>
            </div>
        );
    }
}
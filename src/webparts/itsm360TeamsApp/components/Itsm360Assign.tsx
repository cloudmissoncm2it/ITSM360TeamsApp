import * as React from 'react';
import { Modal, Alert, Button, Icon, Form, Dropdown,Menu } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserDetails } from '../model/IUserDetails';
import { ITeam } from '../model/ITeam';

export interface IItsm360AssignProps {
    visible: boolean;
    sharepointservice: sharepointservice;
    selectedTicket?: ITicketItem;
    ppcontext: WebPartContext;
    selectedRowKeys?: string[];
    teams?: ITeam[];
}

export interface IItsm360AssignState {
    ismodalvisible: boolean;
    modalsave?: boolean;
    assignpersonid?: string;
    errorMessage?: boolean;
    assignedTeam?: any;
    assignedto?:string;
}

export class Itsm360Assign extends React.Component<IItsm360AssignProps, IItsm360AssignState>{

    constructor(props: IItsm360AssignProps) {
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false,
            errorMessage: false
        };
    }

    public handleOk = (e) => {
        const { assignpersonid,assignedTeam } = this.state;
        this.setState({ modalsave: true });
        if (typeof assignpersonid != "undefined" && assignpersonid != "0"&& typeof assignedTeam!="undefined" && assignedTeam.length>0) {
            const updateobj={
                'AssignedPersonId':assignpersonid,
                'AssignedTeamId':assignedTeam
            };
            this.props.selectedRowKeys.forEach((ticketid)=>{
                this.props.sharepointservice.updateTicketAssign(ticketid, updateobj).then((data) => {
                    console.log(`Update the ticket ${ticketid} with the response ${data}`);
                });
            });
            this.setState({
                modalsave: false,
                ismodalvisible: false,
                errorMessage: false
            });

        }else if(typeof assignedTeam!="undefined" && assignedTeam.length>0){
            const updateobj={
                'AssignedTeamId':assignedTeam
            };
            this.props.selectedRowKeys.forEach((ticketid)=>{
                this.props.sharepointservice.updateTicketAssign(ticketid, updateobj).then((data) => {
                    console.log(`Update the ticket ${ticketid} with the response ${data}`);
                });
            });
            this.setState({
                modalsave: false,
                ismodalvisible: false,
                errorMessage: false
            });
        }else if(typeof assignpersonid != "undefined" && assignpersonid != "0"){
            const updateobj={
                'AssignedPersonId':assignpersonid
            };
            this.props.selectedRowKeys.forEach((ticketid)=>{
                this.props.sharepointservice.updateTicketAssign(ticketid, updateobj).then((data) => {
                    console.log(`Update the ticket ${ticketid} with the response ${data}`);
                });
            });
            this.setState({
                modalsave: false,
                ismodalvisible: false,
                errorMessage: false
            });
        }
        else {
            this.setState({
                errorMessage: true,
                modalsave: false
            });
        }

    }

    public handleCancel = (e) => {
        this.setState({ ismodalvisible: false,assignedto:null });
    }

    public _getAssignePickerItems = (people: any[]) => {
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
        let assignedpersonid: IUserDetails[] = [];
        if (people.length > 0) {
            assignedpersonid = this.props.sharepointservice._lusers.filter(i => i.Email == people[0].secondaryText);
        }
        this.setState({ assignpersonid: assignedpersonid.length > 0 ? assignedpersonid[0].ID : "0" });
    }

    public ateamchange = (e) => {
        this.setState({ assignedTeam: e.currentTarget.value });
    }

    public addclick=(e)=>{

        if(e.key=="AM"){
        const Currentusers: IUserDetails[] = this.props.sharepointservice._lusers.filter(i => i.Email == this.props.sharepointservice._currentuser.email);
        const curruser: IUserDetails = Currentusers.length > 0 ? Currentusers[0] : null;
        this.setState({assignedto:curruser.Email,assignpersonid:curruser.ID});
    }
        this.setState({ ismodalvisible: true });
    }

    public render(): React.ReactElement<IItsm360AssignProps> {
        const menu = (
            <Menu onClick={this.addclick}>
                <Menu.Item key="ATP">Assign to Team/Person</Menu.Item>
                <Menu.Item key="AM">Assign to Me</Menu.Item>
            </Menu>
          );

        return (
            <div className="btnattach">
                {/* <Button disabled={!this.props.visible} onClick={() => this.setState({ ismodalvisible: true })}>
                    <Icon type="user-add" />
                    Assign/Re-Assign
                </Button> */}
                <Dropdown overlay={menu} disabled={!this.props.visible} >
                    <Button>
                        Actions <Icon type="down" />
                    </Button>
                </Dropdown>
                <Modal title="Assign/Re-Assign"
                    visible={this.state.ismodalvisible}
                    onOk={this.handleOk}
                    onCancel={this.handleCancel}
                    okText="Update Assign"
                    confirmLoading={this.state.modalsave}
                    destroyOnClose={true}
                >
                    <div>
                        {this.state.errorMessage ? <Alert type="error" style={{ marginBottom: "10px" }} closable message="No user selected, select the user before submitting." /> : ""}
                        <Alert type="info" style={{ marginBottom: "10px" }}
                            message="Selected Ticket IDs"
                            description={this.props.selectedRowKeys.toString()}
                        />
                        <Form layout="horizontal" labelCol={{ span: 8 }} wrapperCol={{ span: 14 }}>
                            <Form.Item label="Assign/Re-Assign">
                                <PeoplePicker
                                    context={this.props.ppcontext}
                                    titleText=""
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    isRequired={false}
                                    disabled={false}
                                    selectedItems={this._getAssignePickerItems}
                                    showHiddenInUI={false}
                                    defaultSelectedUsers={[this.state.assignedto]}
                                    principalTypes={[PrincipalType.User]}
                                    resolveDelay={200} />
                            </Form.Item>
                            <Form.Item label="Assigned Team">
                                <select defaultValue={this.state.assignedTeam} onChange={this.ateamchange} className="ant-select-selection ant-select-selection--single">
                                    <option value="">-select a value-</option>
                                    {this.props.teams.map((team: ITeam, index) => <option value={team.ID} key={index}>{team.Title}</option>)}
                                </select>
                            </Form.Item>
                        </Form>
                    </div>
                </Modal>
            </div>
        );
    }
}
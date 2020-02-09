import * as React from 'react';
import { sharepointservice } from '../service/sharepointservice';
import { Modal, Table, Form, Input, Select, DatePicker, Button, Menu, Dropdown, Icon } from 'antd';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import * as moment from 'moment';
import { ITeam } from '../model/ITeam';
const { Option } = Select;
const { TextArea } = Input;

export interface IItsm360SubTasksProps {
    sharepointservice: sharepointservice;
    ticketid: string;
    ppcontext: WebPartContext;
    teams?: ITeam[];
}

export interface IItsm360SubTasksState {
    tasks?: any[];
    catalog?:any[];
    isVisible?: boolean;
    tasktitle?: string;
    taskstatus?: string;
    staDate?: string;
    dueDate?: string;
    tskdesc?: string;
    tskAssignedTo?: any;
    tskAssignedTeam?: any;
    tskid?:string;
}

export class Itsm360SubTasks extends React.Component<IItsm360SubTasksProps, IItsm360SubTasksState>{

    constructor(props: IItsm360SubTasksProps) {
        super(props);
        this.state = {
            tasks: [],
            catalog:[],
            isVisible: false,
            tskid:""
        };
    }

    public componentDidMount() {
        this.props.sharepointservice.getSubTasks(this.props.ticketid).then((data) => {
            this.setState({ tasks: data });
        });
        this.props.sharepointservice.getTaskCatalog().then((catalogdata)=>{
            this.setState({catalog:catalogdata});
        });
    }

    public titleclick = (e) => {
        const tskid=e.target.id;
        const tsks:any[]=this.state.tasks.filter(i=>i.key==tskid);
        if(tsks.length>0){
            const tsk=tsks[0];
            const persemail=tsk.AssignedPerson!=null?tsk.AssignedPerson.ID:null;
            const tskteam=tsk.AssignedTeam!=null?tsk.AssignedTeam.ID:null;
            this.setState({
                tasktitle:tsk.Title,
                tskdesc:tsk.Description,
                tskAssignedTo:persemail,
                tskAssignedTeam:tskteam,
                taskstatus:tsk.TaskStatus,
                dueDate:tsk.DueDate,
                isVisible:true,
                tskid:tskid
            });    
        }
    }

    public addclick=(e)=>{
        const catalog:any[]=this.state.catalog.filter(i=>i.key==e.key);
        if(catalog.length>0){
            const cat=catalog[0];
            const persemail=cat.AssignedPerson!=null?cat.AssignedPerson.ID:null;
            const catteam=cat.AssignedTeam!=null?cat.AssignedTeam.ID:null;
            this.setState({
                tasktitle:cat.Title,
                tskdesc:cat.Description,
                tskAssignedTo:persemail,
                tskAssignedTeam:catteam
            });
        }
        this.setState({isVisible:true});
    }

    public onSubmit = () => {
        let assignedpersonid;
        const {tskAssignedTo,tskid}=this.state;
        if(tskAssignedTo!=null){
        assignedpersonid = this.props.sharepointservice._lusers.filter(i => i.Email == tskAssignedTo);
    }
        if(tskid.length>0){
        const updatetsk={
            "__metadata":{"type":"SP.Data.TicketSubtasksListItem"},
            Title:this.state.tasktitle,
            Description:this.state.tskdesc,
            TaskStatus:this.state.taskstatus,
            DueDate:this.state.dueDate,
            AssignedTeamId:this.state.tskAssignedTeam,
            AssignedPersonId:typeof assignedpersonid!="undefined" && assignedpersonid.length>0?assignedpersonid[0].ID:null
        };
        this.props.sharepointservice.updateITSMSubTask(updatetsk,this.state.tskid).then((data)=>{
            this.setState({ isVisible: false,tskid:"" });
        });
    }else{
        const createtsk={
            RelatedTicketsId:this.props.ticketid,
            Title:this.state.tasktitle,
            Description:this.state.tskdesc,
            TaskStatus:this.state.taskstatus,
            DueDate:this.state.dueDate,
            AssignedTeamId:this.state.tskAssignedTeam,
            AssignedPersonId:typeof assignedpersonid!="undefined" && assignedpersonid.length>0?assignedpersonid[0].ID:null
        };
        this.props.sharepointservice.addITSMSubTask(createtsk).then((data)=>{
            this.setState({ isVisible: false,tskid:"" });
        }).catch((ex)=>{
            console.log(ex);
        });
    }
    }

    public onCancel = () => {
        this.setState({ isVisible: false });
    }

    public _getPeoplePickerItems = (people: any[]) => {
        let assignedperson:any[]=[];
        if (typeof people != "undefined" && people.length > 0) {
            assignedperson = this.props.sharepointservice._lusers.filter(i => i.Email == people[0].secondaryText);
            if(assignedperson.length>0){this.setState({tskAssignedTo:assignedperson[0].Email});}
        }
    }

    public ateamchange = (e) => {
        this.setState({ tskAssignedTeam: e.currentTarget.value });
    }

    public titleChange = (e) => {
        this.setState({ tasktitle: e.currentTarget.value });
    }

    public descriptionChange = (e) => {
        this.setState({ tskdesc: e.currentTarget.value });
    }

    public onduedatechange=(date,datestring)=>{
        this.setState({dueDate:date.format('YYYY-MM-DDTHH:mm:ssZ')});
    }

    public onstatchange=(value)=>{
        this.setState({taskstatus:value});
    }

    public handleMenuClick=(e)=>{

    }

    public render(): React.ReactElement<IItsm360SubTasksProps> {
        const columns = [
            {
                title: 'Task Name',
                dataIndex: 'Title',
                key: 'Title',
                render: (title, record) => <a id={record.key} onClick={this.titleclick}>{title}</a>
            },
            {
                title: 'Assigned To',
                dataIndex: 'AssignedPerson',
                key: 'AssignedPerson',
                render: ap => <span>{ap != null ? ap.Title : ''}</span>
            },
            {
                title: 'Assigned Team',
                dataIndex: 'AssignedTeam',
                key: 'AssignedTeam',
                render: at => <span>{at != null ? at.Title : ''}</span>
            },
            {
                title: 'Task Status',
                dataIndex: 'TaskStatus',
                key: 'TaskStatus'
            }
        ];
        const dateFormat = 'YYYY/MM/DD';
        const menu = (
            <Menu onClick={this.addclick}>
                {this.state.catalog.map((cat:any,index)=><Menu.Item key={cat.key}>{cat.Title}</Menu.Item>)}
            </Menu>
          );

        return (
            <div>
                <Dropdown overlay={menu} >
                    <Button style={{marginBottom:"10px"}}>
                        Actions <Icon type="down" />
                    </Button>
                </Dropdown>
                <Table columns={columns} dataSource={this.state.tasks} />
                <Modal
                    visible={this.state.isVisible}
                    okText={this.state.tskid.length>0?"Update":"Create"}
                    onOk={this.onSubmit}
                    onCancel={this.onCancel}
                    title="Edit Form"
                    destroyOnClose={true}
                >
                    <Form layout="vertical">
                        <Form.Item label="Task Title">
                            <Input value={this.state.tasktitle} onChange={this.titleChange} />
                        </Form.Item>
                        <Form.Item label="Description">
                            <TextArea value={this.state.tskdesc} onChange={this.descriptionChange} />
                        </Form.Item>
                        <Form.Item label="Assigned To">
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
                                defaultSelectedUsers={[this.state.tskAssignedTo]}
                                resolveDelay={200} />
                        </Form.Item>
                        <Form.Item label="Assigned Team">
                            <select defaultValue={this.state.tskAssignedTeam} onChange={this.ateamchange} className="ant-select-selection ant-select-selection--single">
                            <option value="">-select a value-</option>
                            {this.props.teams.map((team: ITeam, index) => <option value={team.ID} key={index}>{team.Title}</option>)}            
                            </select>
                        </Form.Item>
                        <Form.Item label="Task Status">
                            <Select defaultValue={this.state.taskstatus} onChange={this.onstatchange}>
                                <Option value="Not Started">Not Started</Option>
                                <Option value="In Progress">In Progress</Option>
                                <Option value="Completed">Completed</Option>
                                <Option value="Deferred">Deferred</Option>
                                <Option value="Waiting on someone else">Waiting on someone else</Option>
                            </Select>
                        </Form.Item>
                        {/* <Form.Item label="Start Date">
                            <DatePicker defaultValue={moment(this.state.staDate)} format={dateFormat} />
                        </Form.Item> */}
                        <Form.Item label="Due Date">
                            <DatePicker defaultValue={moment(this.state.dueDate)} format={dateFormat} onChange={this.onduedatechange} />
                        </Form.Item>
                    </Form>
                </Modal>
            </div>
        );
    }
}
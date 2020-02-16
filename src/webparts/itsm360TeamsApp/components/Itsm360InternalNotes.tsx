import * as React from 'react';
import { Form, Input, List, Comment, Checkbox } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import * as moment from 'moment';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { WebPartContext } from '@microsoft/sp-webpart-base';
const { TextArea } = Input;

export interface IItsm360InternalNotesProps {
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
    getinternalnotevalues?: any;
    ppcontext: WebPartContext;
}

export interface IItsm360InternalNotesState {
    internalnotes?: any[];
    newinternalnote?: string;
    isAssignedPerson?: boolean;
    isAssignedTeam?: boolean;
    isAssignedStaff?: boolean;
    staffusers?: any[];
    ppstate?: string;
}

export class Itsm360InternalNotes extends React.Component<IItsm360InternalNotesProps, IItsm360InternalNotesState>{
    constructor(props: IItsm360InternalNotesProps) {
        super(props);
        this.state = {
            internalnotes: [],
            staffusers: [],
            newinternalnote: "",
            isAssignedPerson: false,
            isAssignedTeam: false,
            isAssignedStaff: false,
            ppstate: ""
        };
    }

    public componentDidMount() {
        const spservice = this.props.sharepointservice;
        spservice.getTicketInternalNotes(this.props.selectedTicket.ID).then((notesdata) => {
            this.setState({ internalnotes: notesdata });
        });
    }

    public ticketnoteChange = (e) => {
        this.setState({ newinternalnote: e.currentTarget.value });
        const internalnotevalue = {
            newinternalnote: e.currentTarget.value,
            isAssignedPerson: this.state.isAssignedPerson,
            isAssignedTeam: this.state.isAssignedTeam,
            isAssignedStaff: this.state.isAssignedStaff,
            staffusers: this.state.staffusers
        };
        this.props.getinternalnotevalues(internalnotevalue);
    }

    public assignedpersoncheck = (e) => {
        this.setState({ isAssignedPerson: e.target.checked });
        const internalnotevalue = {
            newinternalnote: this.state.newinternalnote,
            isAssignedPerson: e.target.checked,
            isAssignedTeam: this.state.isAssignedTeam,
            isAssignedStaff: this.state.isAssignedStaff,
            staffusers: this.state.staffusers
        };
        this.props.getinternalnotevalues(internalnotevalue);
    }

    public assignedTeamcheck = (e) => {
        this.setState({ isAssignedTeam: e.target.checked });
        const internalnotevalue = {
            newinternalnote: this.state.newinternalnote,
            isAssignedPerson: this.state.isAssignedPerson,
            isAssignedTeam: e.target.checked,
            isAssignedStaff: this.state.isAssignedStaff,
            staffusers: this.state.staffusers
        };
        this.props.getinternalnotevalues(internalnotevalue);
    }

    public assignedstaffcheck = (e) => {
        this.setState({ isAssignedStaff: e.target.checked });
        const internalnotevalue = {
            newinternalnote: this.state.newinternalnote,
            isAssignedPerson: this.state.isAssignedPerson,
            isAssignedTeam: this.state.isAssignedTeam,
            isAssignedStaff: e.target.checked,
            staffusers: this.state.staffusers
        };
        this.props.getinternalnotevalues(internalnotevalue);
    }

    public _getPeoplePickerItems = (people: any[]) => {
        this.setState({ staffusers: people });
        const internalnotevalue = {
            newinternalnote: this.state.newinternalnote,
            isAssignedPerson: this.state.isAssignedPerson,
            isAssignedTeam: this.state.isAssignedTeam,
            isAssignedStaff: this.state.isAssignedStaff,
            staffusers: people
        };
        this.props.getinternalnotevalues(internalnotevalue);
    }

    public render(): React.ReactElement<IItsm360InternalNotesProps> {
        return (
            <div>
                <Form.Item>
                    <TextArea placeholder="Post a Note" rows={3} style={{ width: "60%" }} onChange={this.ticketnoteChange} value={this.state.newinternalnote} />
                </Form.Item>
                <Form.Item>
                    <Checkbox onChange={this.assignedpersoncheck}>Inform assigned person</Checkbox>
                    <Checkbox onChange={this.assignedTeamcheck}>Inform assigned team</Checkbox>
                    <Checkbox onChange={this.assignedstaffcheck}>Inform other staff members</Checkbox>
                </Form.Item>
                {this.state.isAssignedStaff ? <Form.Item style={{ width: "60%" }}>
                    <PeoplePicker
                        context={this.props.ppcontext}
                        titleText=""
                        personSelectionLimit={6}
                        showtooltip={true}
                        isRequired={false}
                        disabled={false}
                        selectedItems={this._getPeoplePickerItems}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={200} />
                </Form.Item> : null}
                <div ></div>
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
                                content={<div dangerouslySetInnerHTML={{__html: item.content}} />}
                                datetime={moment(item.datetime).fromNow()}
                            />
                        </li>
                    )}
                />
            </div>
        );
    }
}
import * as React from 'react';
import './Itsm360TeamsApp.module.css';
import { Row, Col, Layout, Table, Button, Input, Icon, Statistic, Card, Avatar } from 'antd';
import { ITicketItem } from '../model/ITicketItem';
import { SPHttpClient } from '@microsoft/sp-http';
import { sharepointservice } from '../service/sharepointservice';
import { ISLAPriority } from '../model/ISLAPriority';
import * as microsoftTeams from '@microsoft/teams-js';
import { Istatus } from '../model/Istatus';
import { IContype } from '../model/IContype';
import { SPUser } from '@microsoft/sp-page-context';
import { Itsm360newticket } from './Itsm360newticket';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITeam } from '../model/ITeam';
import { Itsm360StatusUpdate } from './Itsm360StatusUpdate';
import { Itsm360Assign } from './Itsm360Assign';
import { Itsm360Attachment } from './Itsm360Attachment';
import {Itsm360AddNotes} from './Itsm360AddNotes';
import {Itsm360EditTicket} from './Itsm360EditTicket';

export interface IItsm360TeamsAppProps {
  description: string;
  sphttpclient: SPHttpClient;
  currentuser: SPUser;
  context: WebPartContext;
  teamscontext:microsoftTeams.Context;
}

export interface IItsm360TeamsAppState {
  tickets: ITicketItem[];
  priorities: ISLAPriority[];
  statuses: Istatus[];
  conTypes: IContype[];
  teams?: ITeam[];
  pagination: any;
  loading: boolean;
  searchText: string;
  errormessage: string;
  mytickets?: string;
  unassignedtickets?: string;
  opentickets?: string;
  alltickets?: string;
  selectedRowKeys: any[];
  selectedTicket?: ITicketItem;
  isDrawerVisible?:boolean;
  selectectedTicket?:ITicketItem;
}

export class Itsm360TeamsApp extends React.Component<IItsm360TeamsAppProps, IItsm360TeamsAppState> {
  private _spservice: sharepointservice;
  private searchInput;
  private _lastpage:number=10;
  private _nexturl:string;

  constructor(props: IItsm360TeamsAppProps) {
    super(props);
    //this._mockService=new mockdataservice();
    this._spservice = new sharepointservice(this.props.context,this.props.teamscontext);
    this.state = {
      tickets: [],
      priorities: [],
      statuses: [],
      conTypes: [],
      pagination: {},
      loading: false,
      searchText: '',
      errormessage: '',
      selectedRowKeys: [],
      teams: []
    };
  }

  public componentDidMount() {
    this.setState({ loading: true });
    this._spservice.getlookupdata().then((data) => {
      this._spservice.getITSMTickets(null,[],null).then((items) => {
        this.setState({
          tickets: items.tickets,
          priorities: data[1],
          statuses: data[2],
          conTypes: data[0],
          teams: data[3],
          loading: false,
          selectedRowKeys: []
        });
        this._nexturl=items.nexturl;

        this._spservice.getMyTicketsCount().then(mdata => {
          this.setState({ mytickets: mdata });
        });
        this._spservice.getUnassignedTicketsCount().then(undata => {
          this.setState({ unassignedtickets: undata });
        });
        this._spservice.getopenTicketsCount().then(odata => {
          this.setState({ opentickets: odata });
        });
        this._spservice.getallticketscount().then(cdata => {
          debugger;
          this.setState({ alltickets: cdata });
        });
      });
    });


  }

  public refreshticketsdata = (newticket) => {
    this.setState({ loading: true });
    this._spservice.getITSMTickets(null,[],null).then((items) => {
      this.setState({
        loading: false,
        tickets: items.tickets
      });
      this._nexturl=items.nexturl;
    });
    this._spservice.PostToTeams(newticket).then((resp)=>{
      console.log(resp);
    });
  }

  public getColumnSearchProps = dataIndex => ({
    filterDropdown: ({ setSelectedKeys, selectedKeys, confirm, clearFilters }) => (
      <div style={{ padding: 8 }}>
        <Input
          ref={node => {
            this.searchInput = node;
          }}
          placeholder={`Search ${dataIndex}`}
          value={selectedKeys[0]}
          onChange={e => setSelectedKeys(e.target.value ? [e.target.value] : [])}
          onPressEnter={() => this.handleSearch(selectedKeys, confirm)}
          style={{ width: 188, marginBottom: 8, display: 'block' }}
        />
        <Button
          type="primary"
          onClick={() => this.handleSearch(selectedKeys, confirm)}
          icon="search"
          size="small"
          style={{ width: 90, marginRight: 8 }}
        >
          Search
        </Button>
        <Button onClick={() => this.handleReset(clearFilters)} size="small" style={{ width: 90 }}>
          Reset
        </Button>
      </div>
    ),
    filterIcon: filtered => (
      <Icon type="search" style={{ color: filtered ? '#1890ff' : undefined }} />
    ),
  })

  public onSelectChange = selectedRowKeys => {
    console.log('selectedRowKeys changed: ', selectedRowKeys);
    const selectedTicket: ITicketItem = selectedRowKeys.length == 1 ? (this.state.tickets.filter(i => i.ID == selectedRowKeys[0]))[0] : null;
    this.setState({
      selectedRowKeys,
      selectedTicket
    });
  }

  public handleSearch = (selectedKeys, confirm) => {
    debugger;
    confirm();
    this.setState({ searchText: selectedKeys[0] });
  }

  public handleReset = clearFilters => {
    clearFilters();
    this.setState({ searchText: '' });
  }

  public handleTableChange = (pagination, filters, sorter) => {
    const pager = { ...this.state.pagination };
    pager.current = pagination.current;
    if (typeof filters.Priority != "undefined" || typeof filters.Status != "undefined" || typeof filters.ContentType != "undefined" || typeof filters.Title != "undefined") {
      this.setState({loading:true});
      this._spservice.getSearchResults(filters).then((items) => {
        pagination.total = items.length;
        this.setState({
          tickets: items,
          loading: false
        });
      });
    }

    if(pager.current==this._lastpage){
      this.setState({loading:true});
      this._spservice.getITSMTickets(this._nexturl,this.state.tickets,null).then((titems)=>{
        pagination.total=titems.tickets.length;
        this._lastpage=titems.tickets.length/10;
        this._nexturl=titems.nexturl;
        this.setState({
          loading:false,
          tickets:titems.tickets
        });
      });
    }
  }

  public onCardClick=(e)=>{
    let viewname=null;
    if(e.currentTarget.className.indexOf("myview")>-1){
      viewname="myview";
    }else if(e.currentTarget.className.indexOf("unassignedview")>-1){
      viewname="unassignedview";
    }else if(e.currentTarget.className.indexOf("openview")>-1){
      viewname="openview";
    }else if(e.currentTarget.className.indexOf("allview")>-1){
      viewname="allview";
    }

    this.setState({ loading: true });
    this._spservice.getITSMTickets(null,[],viewname).then((items) => {
      this.setState({
        loading: false,
        tickets: items.tickets
      });
      this._nexturl=items.nexturl;
    });
  }

  public onTitleClick=(e)=>{
    debugger;
    this.setState({selectectedTicket:undefined});
    const selectedid=e.currentTarget.id;
    const st=this.state.tickets.filter(i=>i.ID==selectedid);
    this.setState({
      selectectedTicket:st[0]
    });
  }

  public render(): React.ReactElement<IItsm360TeamsAppProps> {
    const { Content, Footer } = Layout;
    const { selectedRowKeys } = this.state;
    const rowSelection = {
      selectedRowKeys,
      onChange: this.onSelectChange,
    };
    const hasSelected = selectedRowKeys.length == 1;

    const columns = [
      // {
      //   title:'ID',
      //   dataIndex:'ID'
      // },
      {
        title: 'Priority',
        dataIndex: 'Priority',
        filters: this.state.priorities.map(item => { return { text: item.Description, value: item.Title }; }),
        render: (text, record) => (
          <div style={{ background: record.Prioritycolor, textAlign: 'center', color: '#ffffff' }}>
            {record.Priority}
          </div>
        ),
        width: '3%'
      },
      {
        title: 'Title',
        dataIndex: 'Title',
        render: (title,record) => <div><Icon type="info-circle" /><Itsm360EditTicket sharepointservice={this._spservice} selectedTicket={record} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} tictitle={title} /></div>,
        ...this.getColumnSearchProps('Title'),
        width: '15%'
      },
      {
        title: 'Requester',
        dataIndex: 'Requester',
        //render:title=><div><Avatar src="https://cloudmission-my.sharepoint.com/User%20Photos/Profilbilleder/kbj_cloudmission_net_MThumb.jpg?t=63671314043" />{title}</div>,
      },
      {
        title: 'Status',
        dataIndex: 'Status',
        filters: this.state.statuses.map(item => { return { text: item.Title, value: item.Title }; })
      },
      {
        title: 'Type',
        dataIndex: 'ContentType',
        filters: this.state.conTypes.map(item => { return { text: item.Name, value: item.ID }; })
      },
      {
        title: 'Assigned Team/Person',
        dataIndex: 'AssignedTeamPerson'
      },
      {
        title: 'Created',
        dataIndex: 'Created',
        width: '10%'
      },
      // {
      //   title: 'Remaining Time',
      //   dataIndex: 'RemainingTime'
      // }
      // {
      //   title:'Action',
      //   render:(text,record)=>(
      //     <span>
      //       <Itsm360EditTicket sharepointservice={this._spservice} selectedTicket={record} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} />
      //     </span>
      //   ),
      // }
    ];

    return (
      <div>
        <Layout className="layout">
          <Content style={{ padding: '0 50px' }}>
            <div className="gutter-example">
              <div style={{ background: '#ECECEC', padding: '30px' }}>
                <Row gutter={16}>
                  <Col className="gutter-row" span={6}>
                    <Card className="myview" onClick={this.onCardClick} style={{cursor:"pointer"}}>
                      <Statistic
                        title="My Tickets"
                        value={this.state.mytickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className="unassignedview" onClick={this.onCardClick} style={{cursor:"pointer"}}>
                      <Statistic
                        title="UnAssigned Tickets"
                        value={this.state.unassignedtickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className="openview" onClick={this.onCardClick} style={{cursor:"pointer"}}>
                      <Statistic
                        title="Open Tickets"
                        value={this.state.opentickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className="allview" onClick={this.onCardClick} style={{cursor:"pointer"}}>
                      <Statistic
                        title="All Tickets"
                        value={this.state.alltickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                </Row>
              </div>
              <Row gutter={16}>
                <Col className="gutter-row" span={24}>

                  <div>
                    <div className="gutter-box">
                      <Itsm360StatusUpdate visible={hasSelected} sharepointservice={this._spservice} selectedTicket={this.state.selectedTicket} status={this.state.statuses} />
                      <Itsm360Assign visible={hasSelected} sharepointservice={this._spservice} selectedTicket={this.state.selectedTicket} ppcontext={this.props.context} />
                      <Itsm360Attachment visible={hasSelected} sharepointservice={this._spservice} selectedTicket={this.state.selectedTicket} />
                      <Itsm360AddNotes sharepointservice={this._spservice} visible={hasSelected} selectedTicket={this.state.selectedTicket} />
                      <Itsm360newticket sharepointservice={this._spservice} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} refreshticketsdata={this.refreshticketsdata} />
                    </div>
                  </div>
                </Col>
              </Row>
              <Row gutter={16}>
                <Col className="gutter-row" span={24}>
                  <div className="gutter-box">

                    <Table dataSource={this.state.tickets} columns={columns} pagination={this.state.pagination} loading={this.state.loading} onChange={this.handleTableChange} rowSelection={rowSelection} />
                  </div>
                </Col>
              </Row>
            </div>
          </Content>
          <Footer style={{ textAlign: 'center' }}>Teams Apps desgined by ITSM 360</Footer>
        </Layout>
      </div>
    );
  }
}

import * as React from 'react';
import './Itsm360TeamsApp.module.css';
import { Row, Col, Layout, Table, Button, Input, Icon, Statistic, Card, Avatar, Modal, Form } from 'antd';
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
import { Itsm360AddNotes } from './Itsm360AddNotes';
import { Itsm360EditTicket } from './Itsm360EditTicket';

export interface IItsm360TeamsAppProps {
  sphttpclient: SPHttpClient;
  currentuser: SPUser;
  context: WebPartContext;
  teamscontext: microsoftTeams.Context;
  spservice: sharepointservice;
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
  isDrawerVisible?: boolean;
  selectectedTicket?: ITicketItem;
  cardselected?: string;
  drawervisible?: boolean;
  siteurl?: string;
  modelloading?: boolean;
}

export class Itsm360TeamsApp extends React.Component<IItsm360TeamsAppProps, IItsm360TeamsAppState> {
  private searchInput;
  private _lastpage: number = 10;
  private _nexturl: string;

  constructor(props: IItsm360TeamsAppProps) {
    super(props);

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
      teams: [],
      drawervisible: false,
      modelloading: false
    };
  }

  public componentDidMount() {
    this.setState({ loading: true });
    this.props.spservice.getlookupdata().then((data) => {
      this.props.spservice.getITSMTickets(null, [], null).then((items) => {
        this.setState({
          tickets: items.tickets,
          priorities: data[1],
          statuses: data[2],
          conTypes: data[0],
          teams: data[3],
          loading: false,
          selectedRowKeys: []
        });
        this._nexturl = items.nexturl;

        this.props.spservice.getMyTicketsCount().then(mdata => {
          this.setState({ mytickets: mdata });
        });
        this.props.spservice.getUnassignedTicketsCount().then(undata => {
          this.setState({ unassignedtickets: undata });
        });
        this.props.spservice.getopenTicketsCount().then(odata => {
          this.setState({ opentickets: odata });
        });
        this.props.spservice.getallticketscount().then(cdata => {
          this.setState({
            alltickets: cdata,
            siteurl: this.props.spservice._weburl
          });
        });
      });
    });


  }

  public refreshticketsdata = (newticket) => {
    debugger;
    this.setState({ loading: true });
    this.props.spservice.getITSMTickets(null, [], null).then((items) => {
      this.setState({
        loading: false,
        tickets: items.tickets
      });
      this._nexturl = items.nexturl;
    });
    if (typeof newticket != "undefined") {
      this.props.spservice.PostToTeams(newticket).then((resp) => {
        console.log(resp);
      });
    }
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
    if (typeof filters.Priority != "undefined" || typeof filters.Status != "undefined" || typeof filters.ContentType != "undefined" || typeof filters.Title != "undefined" || typeof filters.ID != "undefined") {
      this.setState({ loading: true });
      this.props.spservice.getSearchResults(filters).then((items) => {
        pagination.total = items.length;
        this.setState({
          tickets: items,
          loading: false
        });
      });
    }

    if (pager.current == this._lastpage) {
      this.setState({ loading: true });
      this.props.spservice.getITSMTickets(this._nexturl, this.state.tickets, null).then((titems) => {
        pagination.total = titems.tickets.length;
        this._lastpage = titems.tickets.length / 10;
        this._nexturl = titems.nexturl;
        this.setState({
          loading: false,
          tickets: titems.tickets
        });
      });
    }
  }

  public onCardClick = (e) => {
    debugger;
    let viewname = null;
    if (e.currentTarget.id.indexOf("myview") > -1) {
      viewname = "myview";
      this.setState({ cardselected: "1" });
    } else if (e.currentTarget.id.indexOf("unassignedview") > -1) {
      viewname = "unassignedview";
      this.setState({ cardselected: "2" });
    } else if (e.currentTarget.id.indexOf("openview") > -1) {
      viewname = "openview";
      this.setState({ cardselected: "3" });
    } else if (e.currentTarget.id.indexOf("allview") > -1) {
      viewname = "allview";
      this.setState({ cardselected: "4" });
    }

    this.setState({ loading: true });
    this.props.spservice.getITSMTickets(null, [], viewname).then((items) => {
      this.setState({
        loading: false,
        tickets: items.tickets
      });
      this._nexturl = items.nexturl;
    });
  }

  public onTitleClick = (e) => {
    debugger;
    this.setState({ selectectedTicket: undefined });
    const selectedid = e.currentTarget.id;
    const st = this.state.tickets.filter(i => i.ID == selectedid);
    this.setState({
      selectectedTicket: st[0]
    });
  }

  public handleClose = (e) => {
    this.setState({ drawervisible: false });
  }

  public getsitelists = (e) => {
    this.setState({ modelloading: true });
    this.props.spservice.getSiteLists(this.state.siteurl).then((data) => {
      this.setState({
        drawervisible: false,
        modelloading: false
      });
      this.refreshticketsdata(undefined);
    });
  }

  public onSiteurlChange = (e) => {
    this.setState({ siteurl: e.currentTarget.value });
  }

  public onConfigClick=(e)=>{
    this.setState({drawervisible:true});
  }



  public render(): React.ReactElement<IItsm360TeamsAppProps> {
    const { Content, Footer } = Layout;
    const { selectedRowKeys } = this.state;
    const rowSelection = {
      selectedRowKeys,
      onChange: this.onSelectChange,
    };
    const hasSelected = selectedRowKeys.length != 0;

    const columns = [
      {
        title: 'ID',
        dataIndex: 'ID',
        sorter: (a, b) => a.ID - b.ID,
        ...this.getColumnSearchProps('ID')
      },
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
        render: (title, record) => <div><Itsm360EditTicket sharepointservice={this.props.spservice} selectedTicket={record} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} tictitle={title} /></div>,
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
      }
    ];

    return (
      <div>
        <div style={{ position: "relative" }}><Button shape="circle" style={{ position: "absolute", top: "0", right: "0", margin: "5px" }} onClick={this.onConfigClick} ><Icon type="setting" theme="filled" /></Button></div>
        <Layout className="layout">
          <Content style={{ padding: '0 50px' }}>

            <div className="gutter-example">
              <div style={{ background: '#ECECEC', padding: '30px' }}>
                <Row gutter={16}>
                  <Col className="gutter-row" span={6}>
                    <Card className={this.state.cardselected == "1" ? "cardselected" : "cardunselected"} id="myview" onClick={this.onCardClick} >
                      <Statistic
                        title="My Tickets"
                        value={this.state.mytickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className={this.state.cardselected == "2" ? "cardselected" : "cardunselected"} id="unassignedview" onClick={this.onCardClick}>
                      <Statistic
                        title="UnAssigned Tickets"
                        value={this.state.unassignedtickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className={this.state.cardselected == "3" ? "cardselected" : "cardunselected"} id="openview" onClick={this.onCardClick}>
                      <Statistic
                        title="Open Tickets"
                        value={this.state.opentickets}
                        precision={0}
                        valueStyle={{ color: '#3f8600' }}
                      />
                    </Card>
                  </Col>
                  <Col className="gutter-row" span={6}>
                    <Card className={this.state.cardselected == "4" ? "cardselected" : "cardunselected"} id="allview" onClick={this.onCardClick}>
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
                      <Itsm360StatusUpdate visible={hasSelected} sharepointservice={this.props.spservice} selectedRowKeys={this.state.selectedRowKeys} status={this.state.statuses} />
                      <Itsm360Assign visible={hasSelected} sharepointservice={this.props.spservice} selectedRowKeys={this.state.selectedRowKeys}  ppcontext={this.props.context} teams={this.state.teams} />
                      {/* <Itsm360Attachment visible={hasSelected} sharepointservice={this.props.spservice} selectedTicket={this.state.selectedTicket} /> */}
                      <Itsm360AddNotes sharepointservice={this.props.spservice} visible={hasSelected} selectedRowKeys={this.state.selectedRowKeys} />
                      <Itsm360newticket sharepointservice={this.props.spservice} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} refreshticketsdata={this.refreshticketsdata} />
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

        <Modal title="ITSM Configuration" visible={this.state.drawervisible}
          onOk={this.getsitelists} onCancel={this.handleClose}
          confirmLoading={this.state.modelloading}
        >
          <div>
            <Form.Item label="ITSM site url">
              <Input placeholder="ITSM site url" onChange={this.onSiteurlChange} value={this.state.siteurl} />
            </Form.Item>
          </div>
        </Modal>
      </div>
    );
  }
}

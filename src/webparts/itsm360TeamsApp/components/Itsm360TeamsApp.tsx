import * as React from 'react';
import './Itsm360TeamsApp.module.css';
import { Row, Col, Layout, Table, Button, Input, Icon,Statistic,Card } from 'antd';
import { ITicketItem } from '../model/ITicketItem';
import { SPHttpClient } from '@microsoft/sp-http';
import { sharepointservice } from '../service/sharepointservice';
import { ISLAPriority } from '../model/ISLAPriority';
import { Istatus } from '../model/Istatus';
import { IContype } from '../model/IContype';
import { SPUser } from '@microsoft/sp-page-context';
import {Itsm360buttons} from './Itsm360buttons';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITeam } from '../model/ITeam';

export interface IItsm360TeamsAppProps {
  description: string;
  sphttpclient:SPHttpClient;
  currentuser:SPUser;
  context:WebPartContext;
}

export interface IItsm360TeamsAppState{
  tickets:ITicketItem[];
  priorities:ISLAPriority[];
  statuses:Istatus[];
  conTypes:IContype[];
  teams?:ITeam[];
  pagination: any;
  loading: boolean;
  searchText:string;
  errormessage:string;
  mytickets?:string;
  unassignedtickets?:string;
  opentickets?:string;
  alltickets?:string;
  selectedRowKeys:any[];
}

export class Itsm360TeamsApp extends React.Component<IItsm360TeamsAppProps, IItsm360TeamsAppState> {
  private _spservice:sharepointservice;
  private searchInput;

  constructor(props:IItsm360TeamsAppProps){
    super(props);
    //this._mockService=new mockdataservice();
    this._spservice=new sharepointservice(this.props.sphttpclient,this.props.currentuser);
    this.state={
      tickets:[],
      priorities:[],
      statuses:[],
      conTypes:[],
      pagination: {},
      loading: false,
      searchText:'',
      errormessage:'',
      selectedRowKeys:[],
      teams:[]
    };
  }
  
  public componentDidMount(){
    this.setState({loading:true});
    this._spservice.getlookupdata().then((data)=>{
      this._spservice.getITSMTickets().then((items)=>{
        this.setState({
          tickets:items,
          priorities:data[1],
          statuses:data[2],
          conTypes:data[0],
          teams:data[3],
          loading:false,
          selectedRowKeys:[]
        });

        this._spservice.getMyTicketsCount().then(mdata=>{
         this.setState({mytickets:mdata});
        });
        this._spservice.getUnassignedTicketsCount().then(undata=>{
          this.setState({unassignedtickets:undata});
        });
        this._spservice.getopenTicketsCount().then(odata=>{
          this.setState({opentickets:odata});
        });
        this._spservice.getallticketscount().then(cdata=>{
          this.setState({alltickets:cdata});
        });
      });
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
    this.setState({ selectedRowKeys });
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
    this._spservice.getSearchResults(filters).then((items)=>{
      pagination.total=items.length;
      this.setState({
        tickets:items,
        loading:false
      });
    });
  }
  
  public render(): React.ReactElement<IItsm360TeamsAppProps> {
    const { Header, Content, Footer } = Layout;
    const {selectedRowKeys } = this.state;
    const rowSelection = {
      selectedRowKeys,
      onChange: this.onSelectChange,
    };
    const hasSelected = selectedRowKeys.length == 1;

    const columns=[
      // {
      //   title:'ID',
      //   dataIndex:'ID'
      // },
      {
        title:'Priority',
        dataIndex:'Priority',
        filters:this.state.priorities.map(item=>{return {text:item.Description,value:item.Title};}),
        render:(text,record) =>(
         <div style={{background:record.Prioritycolor,textAlign:'center',color:'#ffffff'}}>
           {record.Priority}
         </div> 
        ),
        width:'3%'
      },
      {
        title:'Title',
        dataIndex:'Title',
        render:title=><div><Icon type="info-circle" style={{margin:'0 8px 0 0'}} />{title}</div>,
        ...this.getColumnSearchProps('Title'),
        width:'25%'
      },
      {
        title:'Requester',
        dataIndex:'Requester'
      },
      {
        title:'Status',
        dataIndex:'Status',
        filters:this.state.statuses.map(item=>{return {text:item.Title,value:item.Title};})
      },
      {
        title:'Type',
        dataIndex:'ContentType',
        filters:this.state.conTypes.map(item=>{return {text:item.Name,value:item.ID};})
      },
      {
        title:'Assigned Team/Person',
        dataIndex:'AssignedTeamPerson'
      },
      {
        title:'Created',
        dataIndex:'Created',
        width:'10%'
      },
      {
        title:'Remaining Time',
        dataIndex:'RemainingTime'
      }
    ];

    return (
      <div>
       <Layout className="layout">
       <Content style={{ padding: '0 50px' }}>
            <div className="gutter-example">
              <div style={{ background: '#ECECEC', padding: '30px' }}>
              <Row gutter={16}>
              <Col className="gutter-row" span={6}>
                  <Card>
                    <Statistic
                      title="My Tickets"
                      value={this.state.mytickets}
                      precision={0}
                      valueStyle={{ color: '#3f8600' }}
                    />
                  </Card>
              </Col>
              <Col className="gutter-row" span={6}>
              <Card>
                    <Statistic
                      title="UnAssigned Tickets"
                      value={this.state.unassignedtickets}
                      precision={0}
                      valueStyle={{ color: '#3f8600' }}
                    />
                  </Card>
              </Col>
              <Col className="gutter-row" span={6}>
              <Card>
                    <Statistic
                      title="Open Tickets"
                      value={this.state.opentickets}
                      precision={0}
                      valueStyle={{ color: '#3f8600' }}
                    />
                  </Card>
              </Col>
              <Col className="gutter-row" span={6}>
              <Card>
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
                  <Itsm360buttons hasSelected={hasSelected} sharepointservice={this._spservice} ppcontext={this.props.context} teams={this.state.teams} status={this.state.statuses} />
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
          <Footer style={{ textAlign: 'center' }}>Teams Apps desgined by Thiru</Footer>
       </Layout>
      </div>
    );
  }
}

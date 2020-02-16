import * as React from 'react';
import { Form, Cascader, Select, Row, Col, Table,Button,Descriptions } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { ICI } from '../model/ICI';
import { IAsset } from '../model/IAsset';
const { Option } = Select;

export interface IItsm360ClassificationProps {
    sharepointservice: sharepointservice;
    selectedTicket: ITicketItem;
    defaultvalues: any;
    getclassificationvalues?: any;
}

export interface IItsm360ClassificationState {
    cascaderoptions?: any[];
    servicegroup?: string;
    service?: string;
    subcategory?: string;
    CIs?: ICI[];
    selectedCI?: string;
    assets?: IAsset[];
    selectedRowKeys?: any[];
    loading?:boolean;
}

export class Itsm360Classification extends React.Component<IItsm360ClassificationProps, IItsm360ClassificationState>{
    private _sgtitle:string;
    private _setitle:string;
    private _scategorytitle:string;
    private _relatedCi:string;

    constructor(props: IItsm360ClassificationProps) {
        super(props);
        this.state = {
            cascaderoptions: [],
            CIs: [],
            assets: [],
            selectedRowKeys: [],
            loading:false
        };
    }

    public componentDidMount() {
        const { servicegroupid, serviceid, subcategoryid, relatedCiid,servicegrouptitle,servicetitle,subcategorytitle,relatedCititle,relatedassets } = this.props.defaultvalues;
        this._sgtitle=servicegrouptitle;
        this._setitle=servicetitle;
        this._scategorytitle=subcategorytitle;
        this._relatedCi=relatedCititle;
        const spservice = this.props.sharepointservice;
        this.setState({selectedRowKeys:relatedassets});
        spservice.getlookupdatanew().then((cdata) => {
            this.setState({
                cascaderoptions: cdata,
                servicegroup: servicegroupid,
                service: serviceid,
                subcategory: subcategoryid
            });
        });
        spservice.getCIsLookUp().then((cidata) => {
            this.setState({
                CIs: cidata,
                selectedCI: relatedCiid
            });
        });

        spservice.getuserAssets(this.props.selectedTicket).then((adata) => {
            this.setState({ assets: adata });
        });
    }

    public cascaderChange = (e) => {
        let clasificationvalue;
        if (e.length == 3) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: e[2]
            });
            clasificationvalue={
                servicegroupid: e[0],
                serviceid: e[1],
                subcategoryid: e[2],
                relatedCiid: this.state.selectedCI
            };
        } else if (e.length == 2) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: "-1"
            });
            clasificationvalue={
                servicegroupid: e[0],
                serviceid: e[1],
                subcategoryid: "-1",
                relatedCiid: this.state.selectedCI
            };
        } else if (e.length == 1) {
            this.setState({
                servicegroup: e[0],
                service: "-1",
                subcategory: "-1"
            });
            clasificationvalue={
                servicegroupid: e[0],
                serviceid: "-1",
                subcategoryid: "-1",
                relatedCiid: this.state.selectedCI
            };
        } else {
            this.setState({
                servicegroup: "-1",
                service: "-1",
                subcategory: "-1"
            });
            clasificationvalue={
                servicegroupid: "-1",
                serviceid: "-1",
                subcategoryid: "-1",
                relatedCiid: this.state.selectedCI
            };
        }

        this.props.getclassificationvalues(clasificationvalue);
    }

    public cichange = (e) => {
        this.setState({ selectedCI: e });
        const clasificationvalue = {
            servicegroupid: this.state.servicegroup,
            serviceid: this.state.service,
            subcategoryid: this.state.subcategory,
            relatedCiid: e
        };

        this.props.getclassificationvalues(clasificationvalue);
    }

    public onSelectChange = selectedRowKeys => {
        //console.log('selectedRowKeys changed: ', selectedRowKeys);
        this.setState({ selectedRowKeys });
    }

    public addrelatedassets=(e)=>{
        const {selectedRowKeys}=this.state;
        if(selectedRowKeys.length>0){
        this.setState({loading:true});
        const addasset={
            "__metadata":{"type":"SP.Data.TicketsListItem"},
            'RelatedAssetsId':{
                'results':selectedRowKeys
            }
        };
        this.props.sharepointservice.updateTicketRelatedAssets(addasset,this.props.selectedTicket.ID).then((data)=>{
            console.log(data);
            this.setState({loading:false});
        });
    }
    }

    public render(): React.ReactElement<IItsm360ClassificationProps> {
        const columns = [
            {
                title: 'Title',
                dataIndex: 'Title',
            },
            {
                title: 'Model',
                dataIndex: 'Model',
            },
            {
                title: 'Type',
                dataIndex: 'Type',
            },
            {
                title: 'State',
                dataIndex: 'State',
            }
        ];
        const { selectedRowKeys } = this.state;
        const rowSelection = {
            selectedRowKeys,
            onChange: this.onSelectChange,
        };
        return (
            <div>
                <Row align="middle" type="flex" justify="space-between">
                    <Col span={10}>
                        <Form layout="vertical">
                            <Form.Item label="Classification">
                                <Cascader options={this.state.cascaderoptions} onChange={this.cascaderChange} defaultValue={[this.state.servicegroup, this.state.service, this.state.subcategory]} />
                            </Form.Item>

                            <Form.Item label="Related CI">
                                <Select placeholder="Select CIs" onChange={this.cichange} value={this.state.selectedCI}>
                                    {this.state.CIs.map((item: ICI, index) => <Option value={item.ID} key={index}>{item.Title}</Option>)}
                                </Select>
                            </Form.Item>
                        </Form>
                    </Col>
                    <Col span={14}>
                    <Descriptions bordered style={{marginLeft:"1%"}} size="small">
                                 <Descriptions.Item label="Service Group" span={3}>{this._sgtitle}</Descriptions.Item>
                                 <Descriptions.Item label="Service" span={3}>{this._setitle}</Descriptions.Item>
                                 <Descriptions.Item label="Category" span={3}>{this._scategorytitle}</Descriptions.Item>
                                 <Descriptions.Item label="RelatedCI" span={3}>{this._relatedCi}</Descriptions.Item>         
                    </Descriptions>
                    </Col>
                </Row>
                <Row align="middle" type="flex" justify="space-between">
                    <h4>Requestor Assets</h4>
                    <Button type="primary" onClick={this.addrelatedassets} loading={this.state.loading}>Update Related Assets</Button>
                    <Table dataSource={this.state.assets} columns={columns} rowSelection={rowSelection} />
                </Row>
            </div>
        );
    }
}
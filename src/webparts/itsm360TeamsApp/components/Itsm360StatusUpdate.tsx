import * as React from 'react';
import { Modal,Alert,Button,Icon } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';
import { Istatus } from '../model/Istatus';

export interface IItsm360StatusUpdateProps{
    visible:boolean;
    sharepointservice:sharepointservice;
    selectedTicket:ITicketItem;
    status:Istatus[];
}

export interface IItsm360StatusUpdateState{
    ismodalvisible:boolean;
    modalsave?:boolean;
    statusid?:string;
}

export class Itsm360StatusUpdate extends React.Component<IItsm360StatusUpdateProps,IItsm360StatusUpdateState>{

    constructor(props:IItsm360StatusUpdateProps){
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false
        };
    }

    public handleOk=(e)=>{
        this.setState({modalsave:true});
        this.props.sharepointservice.updateTicketStatus(this.props.selectedTicket.ID,this.state.statusid).then((resp)=>{
            console.log("update status: ",resp);
            this.setState({
                modalsave:false,
                ismodalvisible:false
            });
        });
    }

    public handleCancel=(e)=>{
        this.setState({ismodalvisible:false});
    }

    public statusChange = (e) => {
        this.setState({statusid:e.target.value});
      }

    public render():React.ReactElement<IItsm360StatusUpdateProps>{
        const stats:Istatus[]=this.props.status;
        const ticketdesc=this.props.selectedTicket?this.props.selectedTicket.Title:"";
        const ticketid=this.props.selectedTicket?this.props.selectedTicket.ID:"";
        const selectedstatus=this.props.selectedTicket?this.props.selectedTicket.Status:"";

        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={()=> this.setState({ismodalvisible:true})}>
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
                        <Alert type="info" 
                            message={`Title: ${ticketdesc}`}
                            description={`Ticket Id: ${ticketid}`}
                        />
                        <div className="ant-form ant-form-horizontal">
                        <div className="ant-row ant-form-item" >
                    <div className="ant-col ant-form-item-label ant-col-xs-24 ant-col-sm-8">
                      <label>Status</label>
                    </div>
                    <div className="ant-col ant-form-item-control-wrapper ant-col-xs-24 ant-col-sm-16">
                      <div className="ant-form-item-control">
                        <span className="ant-form-item-children">
                        <select onChange={this.statusChange} defaultValue={selectedstatus}>
                        {stats.map((stat:Istatus,index)=>{
                          if(stat.Title==selectedstatus){
                            return <option value={stat.ID} key={index} selected>{stat.Title}</option>;
                          }else{  
                            return <option value={stat.ID} key={index}>{stat.Title}</option>;}
                        })}
                        </select>
                        </span>
                      </div>
                    </div>
                  </div>
                        </div>
                    </div>
                </Modal>
            </div>
        );
    }
}
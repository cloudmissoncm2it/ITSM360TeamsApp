import * as React from 'react';
import './Itsm360TeamsApp.module.css';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sharepointservice } from '../service/sharepointservice';
import {Itsm360TeamsApp} from './Itsm360TeamsApp';
import * as microsoftTeams from '@microsoft/teams-js';
import { Spin, Form, Input,Modal } from 'antd';

export interface IItsm360TeamsConfigureProps{
    context:WebPartContext;
    teamscontext:microsoftTeams.Context;
}

export interface IItsm360TeamsConfigureState {
  configvisible?: boolean;
  initialloading?: boolean;
  cardlayout?: boolean;
  drawervisible?: boolean;
  siteurl?: string;
  lstopts?: any[];
  tktlst?:string;
  stlst?:string;
  sglst?:string;
  modelloading?:boolean;
}

export class Itsm360TeamsConfigure extends React.Component<IItsm360TeamsConfigureProps, IItsm360TeamsConfigureState> {
    private _spservice: sharepointservice;
  
    constructor(props: IItsm360TeamsConfigureProps) {
    super(props);
    this._spservice = new sharepointservice(this.props.context,this.props.teamscontext);
    this.state = {
      configvisible: false,
      initialloading: true,
      modelloading:false,
      lstopts:[]
    };
  }

  public componentDidMount() {
    this._spservice.getStorageEntity().then((msg) => {
      this.setState({
        initialloading: false,
        configvisible: false,
        cardlayout: true,
      });
    }).catch((msg) => {
      this.setState({
        initialloading: false,
        configvisible: true,
        cardlayout: false
      });
    });
  }


  public _onConfigure = () => {
    this.setState({ drawervisible: true });
  }

  public handleClose = (e) => {
    this.setState({ drawervisible: false });
  }

  public getsitelists = (e) => {
    this.setState({modelloading:true});
    this._spservice.getSiteLists(this.state.siteurl).then((data) => {
      this.setState({
        initialloading: false,
        configvisible: false,
        cardlayout: true,
        drawervisible:false,
        modelloading:false
      });
    });
  }

  public onSiteurlChange = (e) => {
    this.setState({ siteurl: e.currentTarget.value });
  }

  public render(): React.ReactElement<IItsm360TeamsConfigureProps> {
    return (
      <div>
        <div>
          {this.state.initialloading ?
            <Spin tip="Getting configuration data">
            </Spin> : ""}
          {this.state.configvisible ? <Placeholder iconName='Edit'
            iconText='Configure ITSM Teams App'
            description='Please configure your ITSM Teams App'
            buttonLabel='Configure'
            onConfigure={this._onConfigure} /> : ""}
          {this.state.cardlayout ? <Itsm360TeamsApp context={this.props.context} teamscontext={this.props.teamscontext} sphttpclient={null} currentuser={null} spservice={this._spservice}  /> : ""}
        </div>
        <div>
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
      </div>
    );
  }
}

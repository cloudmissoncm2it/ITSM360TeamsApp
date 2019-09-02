import * as React from 'react';
import { Modal,Upload,Icon,Alert,Button } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { ITicketItem } from '../model/ITicketItem';

export interface IItsm360AttachmentProps{
    visible:boolean;
    sharepointservice:sharepointservice;
    selectedTicket:ITicketItem;
}

export interface IItsm360AttachmentState{
    ismodalvisible:boolean;
    modalsave?:boolean;
}

export class Itsm360Attachment extends React.Component<IItsm360AttachmentProps,IItsm360AttachmentState>{

    constructor(props:IItsm360AttachmentProps){
        super(props);
        this.state = {
            ismodalvisible: this.props.visible,
            modalsave: false
        };
    }

    

    public handleOk=(e)=>{
        this.setState({ismodalvisible:false});
    }

    public handleCancel=(e)=>{
        this.setState({ismodalvisible:false});
    }

    public uploadfilechange=info=>{
        this.props.sharepointservice.uploadTicketAttachment(info.file,this.props.selectedTicket.ID).then(data=>{
            info.onSuccess(null,info.file);
        });
    }


    public render():React.ReactElement<IItsm360AttachmentProps>{
        const {Dragger}=Upload;
        const props={
            name:"item attachment",
            multiple:false,
            customRequest:this.uploadfilechange
        };

        const ticketdesc=this.props.selectedTicket?this.props.selectedTicket.Title:"";
        const ticketid=this.props.selectedTicket?this.props.selectedTicket.ID:"";

        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={()=> this.setState({ismodalvisible:true})}>Add attachment</Button>
                <Modal title="Add Attachment"
                   visible={this.state.ismodalvisible}
                   onOk={this.handleOk}
                   onCancel={this.handleCancel}
                   okText="Submit"
                   confirmLoading={this.state.modalsave} 
                >
                    <div>
                        <Alert type="info" 
                            message={`Title: ${ticketdesc}`}
                            description={`Ticket Id: ${ticketid}`}
                        />
                        <Dragger {...props}>
                            <p className="ant-upload-drag-icon">
                                <Icon type="inbox" />
                            </p>
                            <p className="ant-upload-text">Click or drag files to this area to upload</p>
                            <p className="ant-upload-hint">
                                Support for a single upload. Add a supported file to the ticket.Sharepoint file limitations apply.
                            </p>
                        </Dragger>
                    </div>
                </Modal>
            </div>
        );
    }
}
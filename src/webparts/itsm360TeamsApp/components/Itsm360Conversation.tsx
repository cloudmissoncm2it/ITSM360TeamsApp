import * as React from 'react';
import { Select, Row, Col, Spin, Button, List,Comment } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
import { ITicketItem } from '../model/ITicketItem';
import { IUserDetails } from '../model/IUserDetails';
import * as moment from 'moment';

const { Option } = Select;

export interface IItsm360ConversationProps {
    spservice: sharepointservice;
    classificationvals?:any;
    selectedTicket:ITicketItem;
}

export interface IItsm360ConversationState {
    newconversation?:any;
    loading?:boolean;
    KBArticles?:any[];
    stdMsgs?:any[];
    matMsgs?:any[];
    docs?:any[];
    notesdata?:any[];
    selectedKBs?:any[];
    selectedDocs?:any[];
}

export class Itsm360Conversation extends React.Component<IItsm360ConversationProps, IItsm360ConversationState>{
    private _count=0;

    constructor(props: IItsm360ConversationProps) {
        super(props);
        this.state = {
            newconversation:"",
            loading:false,
            KBArticles:[],
            stdMsgs:[],
            docs:[],
            matMsgs:[],
            notesdata:[],
            selectedDocs:[],
            selectedKBs:[]
        };
    }

    public componentDidMount(){
        this.setState({loading:true});
        const {spservice,selectedTicket}= this.props;
        spservice.getKBArtciles().then((KBs)=>{
            this.setState({KBArticles:KBs});
        });
        spservice.getsitedocuments().then((docs)=>{
            this.setState({docs:docs});
        });
        spservice.getStandardMessages().then((msgs:any[])=>{
            const {serviceid,subcategoryid}=this.props.classificationvals;
            this.setState({stdMsgs:msgs,loading:false});
            const xmsgs=msgs.filter(i=>i.ServiceId==serviceid || i.CategoryId==subcategoryid);
            this.setState({matMsgs:xmsgs});
        });

        spservice.getTicketNotes(selectedTicket.ID).then((notesdata) => {
            this.setState({ notesdata: notesdata });
        });

    }

    public stdmsgchange=(val)=>{
        debugger;
        const selectedmsg=this.state.stdMsgs.filter(i=>i.ID==val);
        if(selectedmsg.length>0){
            ++this._count;
            this.setState({newconversation:selectedmsg[0].Message});
        }
    }

    public conchange=(e)=>{
        this.setState({newconversation:e});
        return e;
    }

    public postinternalnotes=(e)=>{
        const { newconversation, notesdata,selectedDocs,selectedKBs,docs,KBArticles } = this.state;
        const {spservice,selectedTicket}=this.props;
        if (typeof newconversation != "undefined") {
            const Currentusers: IUserDetails[] = spservice._lusers.filter(i => i.Email == spservice._currentuser.email);
            const user: IUserDetails = Currentusers.length > 0 ? Currentusers[0] : null;
            let ldocs:string="";
            selectedDocs.forEach((docid)=>{
                const xdoc= docs.filter(i=>i.ID==docid);
                if(xdoc.length>0){
                    ldocs=`${ldocs}{"key":${docid},"text":${xdoc[0].Title}},`;
                }
            });
            ldocs=`[${ldocs}]`;
            let kbs={'__metadata': { type: 'Collection(Edm.Int32)' },'results':selectedKBs};
            const tnote = {
                '__metadata': {
                    'type': 'SP.Data.TicketCommunicationsListItem'
                },
                'TicketIDId': selectedTicket.ID,
                'Communications': newconversation,
                'CommunicationInitiatorId': Currentusers.length > 0 ? Currentusers[0].ID : null,
                'AttachedKnowledgeId':kbs,
                'AttachedDocuments':ldocs
            };

            spservice.addTicketNotes(tnote).then((tdata) => {
                spservice.getTicketNotes(selectedTicket.ID).then((ndata) => {
                    this.setState({ notesdata: ndata });
                    ++this._count;
                    this.setState({selectedKBs:[],selectedDocs:[],newconversation:""});
                });
            });
        }
    }

    public render(): React.ReactElement<IItsm360ConversationProps> {

        const {newconversation,loading,KBArticles,stdMsgs,docs,matMsgs,selectedDocs,selectedKBs}=this.state;
        return (
            <div>
                {loading?<Spin />:<div></div>}
                <Row align="middle" type="flex" justify="space-between">
                    <Col span="12">
                        <Select style={{width:"90%"}} onChange={this.stdmsgchange} placeholder="All standard responses">
                        {stdMsgs.map((stdmsg:any, index) => <Option value={stdmsg.ID} key={index}>{stdmsg.Title}</Option>)}  
                        </Select>
                    </Col>
                    <Col span="12">
                        <Select style={{width:"90%"}} onChange={(val:any)=>{this.setState({newconversation:val});}} placeholder="Matching standard responses">
                        {matMsgs.map((stdmsg:any, index) => <Option value={stdmsg.ID} key={index}>{stdmsg.Title}</Option>)} 
                        </Select>
                    </Col>
                </Row>
                <Row align="middle" type="flex" justify="space-between" style={{marginTop:"10px"}}>
                    <Col span="24">
                        <div style={{ border: "1px solid #d9d9d9",maxHeight:"150px",overflow:"auto"}} key={this._count}>
                         <RichText isEditMode={true} value={newconversation} onChange={this.conchange} />   </div>
                    </Col>
                </Row>
                <Row align="middle" type="flex" justify="space-between" style={{marginTop:"10px"}}>
                <Col span="12">
                        <Select mode="multiple" style={{width:"90%"}} value={selectedKBs} onChange={(val:any)=>{this.setState({selectedKBs:val});}} placeholder="Select KB articles">
                        {KBArticles.map((KBart:any, index) => <Option value={KBart.ID} key={index}>{KBart.Title}</Option>)} 
                        </Select>
                    </Col>
                    <Col span="12">
                        <Select mode="multiple" style={{width:"90%"}} value={selectedDocs} onChange={(val:any)=>{this.setState({selectedDocs:val});}} placeholder="Select documents">
                            {docs.map((doc:any, index) => <Option value={doc.ID} key={index}>{doc.Title}</Option>)}
                        </Select>
                    </Col> 
                </Row>
                <Row align="middle" type="flex" justify="space-between" style={{marginTop:"10px"}}>
                    <Col span="20"></Col>
                    <Col span="4">
                        <Button type="primary" icon="message" size="small" onClick={this.postinternalnotes}>Post</Button>
                    </Col>
                </Row>
                <Row align="middle" type="flex" justify="space-between" style={{marginTop:"10px"}}>
                    <Col span="20">
                        <List
                            className="comment-list"
                            header={`${this.state.notesdata.length} replies`}
                            itemLayout="horizontal"
                            dataSource={this.state.notesdata}
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
                    </Col>
                </Row>
            </div>
        );
    }
}


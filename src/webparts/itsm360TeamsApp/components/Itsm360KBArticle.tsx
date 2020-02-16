import * as React from 'react';
import { Modal, Alert, Button, Icon, Form, Input, Cascader, Switch } from 'antd';
import { sharepointservice } from '../service/sharepointservice';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";

export interface IItsm360KBArticleProps {
    sharepointservice: sharepointservice;
    selectedrowkeys?: string[];
    visible: boolean;
}

export interface IItsm360KBArticleState {
    ismodalvisible: boolean;
    modalsave?: boolean;
    errorMessage?: string;
    KBDescription?: string;
    cascaderoptions?: any[];
    langs?:any[];
    servicegroup?: string;
    service?: string;
    subcategory?: string;
    validatestatus?:any;
    KBTitle?:string;
    selectedlang?:string;
    selectedcatg?:string;
    pubendusers?:boolean;
}


export class Itsm360KBArticle extends React.Component<IItsm360KBArticleProps, IItsm360KBArticleState>{
    constructor(props: IItsm360KBArticleProps) {
        super(props);
        this.state={
            ismodalvisible:false,
            cascaderoptions: [],
            langs:[],
            validatestatus:'validating'
        };
    }
    
    public onKBClick=(e)=>{
        const {selectedrowkeys}=this.props;
        this.setState({ ismodalvisible: true });
        this.props.sharepointservice.getlookupdatanew().then((cdata) => {
            this.props.sharepointservice.getTicketDetails(selectedrowkeys[0]).then((cds)=>{
                this.setState({
                    servicegroup:cds.ServiceGroups.ID,
                    service:cds.RelatedServices.ID,
                    subcategory:cds.RelatedCategories.ID,
                    cascaderoptions: cdata,
                    KBDescription:cds.Description
                });
            });
        });

        this.props.sharepointservice.getLanguages().then((langs)=>{
            this.setState({langs});
            langs.forEach(element => {
                if(element.DefaultLanguage){
                    this.setState({selectedlang:element.ID});
                }    
            });
        });
    }

    public handleOk = (e) => {
        const {KBTitle,KBDescription,service,subcategory,selectedcatg,selectedlang,pubendusers}=this.state;
        this.setState({modalsave:true});
        if(typeof KBTitle!="undefined" && KBTitle.length>0){
            const KBObj={
                '__metadata': {
                    'type': 'SP.Data.KnowledgeArticlesListItem'
                },
                'Title':KBTitle,
                'Objective':KBDescription,
                'Category':selectedcatg,
                'RelatedServicesId':{'__metadata': { type: 'Collection(Edm.Int32)' },'results':[Number(service)]},
                'RelatedSubcategoriesId':{'__metadata': { type: 'Collection(Edm.Int32)' },'results':[Number(subcategory)]},
                'LanguageId':selectedlang,
                'PublishForEndUsers':pubendusers
            };
            console.log(JSON.stringify(KBObj));
            this.props.sharepointservice.addKBArticle(KBObj).then((data)=>{
                console.log(data);
            });
            this.setState({ ismodalvisible: false,modalsave:false });
        }else{
            this.setState({
                validatestatus:"error",
                errorMessage:"Title is a mandatory field",
                ismodalvisible:true,
                modalsave:false
            });

        }
        
    }

    public handleCancel = (e) => {
        this.setState({ ismodalvisible: false,KBDescription:"" });
    }

    public descriptionChange = (value: string) => {
        this.setState({ KBDescription: value });
        return value;
    }

    public cascaderChange = (e) => {
        if (e.length == 3) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: e[2]
            });
        } else if (e.length == 2) {
            this.setState({
                servicegroup: e[0],
                service: e[1],
                subcategory: "-1"
            });
        } else if (e.length == 1) {
            this.setState({
                servicegroup: e[0],
                service: "-1",
                subcategory: "-1"
            });
        } else {
            this.setState({
                servicegroup: "-1",
                service: "-1",
                subcategory: "-1"
            });
            
        }
    }

    public titleChange = (e) => {
        this.setState({ KBTitle: e.currentTarget.value });
    }

    public langchange = (e) => {
        this.setState({ selectedlang: e.currentTarget.value });
    }

    public catgchange = (e) => {
        this.setState({ selectedcatg: e.currentTarget.value });
    }

    public render(): React.ReactElement<IItsm360KBArticleProps> {
        const {KBDescription,servicegroup,service,validatestatus,errorMessage,KBTitle,selectedlang,selectedcatg}=this.state;
        return (
            <div className="btnattach">
                <Button disabled={!this.props.visible} onClick={this.onKBClick}>
                    <Icon type="file-add" />
                    Knowledge
                </Button>
                <Modal title="Create KB Article"
                    visible={this.state.ismodalvisible}
                    onOk={this.handleOk}
                    onCancel={this.handleCancel}
                    okText="Create KB Article"
                    confirmLoading={this.state.modalsave}
                    destroyOnClose={true}
                >
                    <Form layout="vertical" labelCol={{ span: 8 }} wrapperCol={{ span: 12 }}>
                        <Form.Item label="Title" validateStatus={validatestatus} help={errorMessage}>
                            <Input placeholder="Title of KB Article" value={KBTitle} onChange={this.titleChange} />
                        </Form.Item>
                        <Form.Item label="Description">{typeof KBDescription!="undefined"?
                        <div style={{ border: "1px solid #d9d9d9" }}>
                        <RichText value={this.state.KBDescription} onChange={this.descriptionChange} isEditMode={true} /></div>
                        :""}
                        </Form.Item>
                        <Form.Item label="Category">
                            <select className="ant-select-selection ant-select-selection--single" value={selectedcatg} onChange={this.catgchange} >
                                <option value="" selected>Select a caetgory</option>
                                <option value="User guide">User guide</option>
                                <option value="Technical guide">Technical guide</option>
                                <option value="FAQ">FAQ</option>
                                <option value="Tutorial">Tutorial</option>
                                <option value="Policy">Policy</option>
                            </select>
                        </Form.Item>
                        <Form.Item label="Publish For End Users">
                            <Switch onChange={(checked)=>this.setState({pubendusers:checked})} />
                        </Form.Item>
                        <Form.Item label="Language">
                            <select className="ant-select-selection ant-select-selection--single" value={selectedlang} onChange={this.langchange}>
                            {this.state.langs.map((lang: any, index) => <option value={lang.ID} key={index}>{lang.Title}</option>)} 
                            </select>
                        </Form.Item>
                        <Form.Item label="Classification">{typeof servicegroup!="undefined"||typeof service !="undefined"?
                            <Cascader options={this.state.cascaderoptions} onChange={this.cascaderChange} defaultValue={[this.state.servicegroup, this.state.service, this.state.subcategory]} />
                            :""}
                        </Form.Item>
                    </Form>
                </Modal>
            </div>);
    }
}
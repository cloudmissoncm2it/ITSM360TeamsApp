import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions, ISPHttpClientBatchOptions, SPHttpClientBatch, ISPHttpClientBatchCreationOptions, HttpClient, IHttpClientOptions, HttpClientResponse } from "@microsoft/sp-http";
import { ITicketItem } from "../model/ITicketItem";
import * as moment from 'moment';
import { ISLAPriority } from "../model/ISLAPriority";
import { Istatus } from "../model/Istatus";
import { IContype } from "../model/IContype";
import * as microsoftTeams from '@microsoft/teams-js';
import { ITeam } from "../model/ITeam";
import { IUserDetails } from "../model/IUserDetails";
import { SPUser } from "@microsoft/sp-page-context";
import { IServiceGroup } from "../model/IServiceGroup";
import { IService } from "../model/IService";
import { IServiceCategory } from "../model/IServiceCategory";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICI } from "../model/ICI";
import { IAsset } from "../model/IAsset";
import { sp} from "@pnp/sp/presets/all";
import { ITask } from "../model/ITask";

export class sharepointservice{
    private _spclient:SPHttpClient;
    private _teamscontext:microsoftTeams.Context;
    private _httpclient:HttpClient;
    public _weburl;
    private _ticketsid;
    private _teamsid;
    private _prioritesid;
    private _statusid;
    private _servicegroupid;
    private _servicesid;
    private _subcategory;
    private _conversationid;
    private _emailsid="d6004eaf-11a3-429c-b6a4-5b76677ea7c5";
    private _CIid;
    private _assetsid;
    private _ticketnotesid;
    private _tickettasksid="f0b6bde3-e5ff-415e-be1e-35c888f27f00";
    private _taskCatalog="2cbc3620-691f-4da5-ae60-8f17f9cbdb69";
    private _languageid="0a0fb546-cb9b-4864-a1c0-690ce27d5ff6";
    private _KBArtcileid="c676b182-0f57-4cac-9475-6c7e41e53632";
    private _standMessage="ab2e4c11-0ba1-4332-8b26-27b20245f16d";
    private _doclibraryid="2926729c-fd33-474a-a254-c2e79fc0a9d7";
    private _spris:ISLAPriority[]=[];
    private _stats:Istatus[]=[];
    private _sconts:IContype[]=[];
    private _steams:ITeam[]=[];
    public _lusers:IUserDetails[]=[];
    public _CIs:ICI[]=[];
    public _currentuser:SPUser;

    constructor(context:WebPartContext,teamscontext:microsoftTeams.Context){
        this._spclient=context.spHttpClient;
        this._currentuser=context.pageContext.user;
        this._teamscontext=teamscontext;
        this._httpclient=context.httpClient;
    }

    public getStorageEntity():Promise<any>{
        return sp.web.getStorageEntity("itsmconfig").then((res)=>{
            if(typeof res.Value!="undefined"){
                const configdata=JSON.parse(res.Value);
                this._ticketsid=configdata.ticketid;
                this._weburl=configdata.siteurl;
                this._statusid=configdata.statusid;
                this._teamsid=configdata.teamsid;
                this._prioritesid=configdata.prioritesid;
                this._servicegroupid=configdata.servicegroupid;
                this._servicesid=configdata.servicesid;
                this._subcategory=configdata.subcategory;
                this._conversationid=configdata.conversationid;
                this._CIid=configdata.CIid;
                this._assetsid=configdata.assetsid;
                this._ticketnotesid=configdata.ticketnotesid;
                this.getUsers(null);
                this.getCIsLookUp().then((data)=>{
                    this._CIs=data;
                });
                return Promise.resolve("Success");
            }else{
                return Promise.reject("Error! No Storage entity found");
            }
        });
    }

    public setStorageEntity(configdata:any):Promise<any>{
        const cdata=JSON.stringify(configdata);
        sp.getTenantAppCatalogWeb().then((catweb)=>{
            catweb.setStorageEntity("itsmconfig",cdata);
        });

        return Promise.resolve("Success");
    }

    public getSiteLists(url:string):Promise<any>{
        const querygetAllLsts = `${url}/_api/web/Lists?$select=Title,Id,RootFolder/Name&$filter=Hidden eq false and BaseTemplate eq 100&$expand=RootFolder`;
        this._weburl=url;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllLsts, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                data.value.forEach((x)=>{
                    switch(x.RootFolder.Name){
                        case 'Tickets':
                            this._ticketsid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'Teams':
                            this._teamsid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'SLAPriorities':
                            this._prioritesid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'TicketStatus':
                            this._statusid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'ServiceGroups':
                            this._servicegroupid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'Services':
                            this._servicesid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'ServiceCategories':
                            this._subcategory=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'TicketCommunications':
                            this._conversationid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'CIs':
                            this._CIid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'Assets':
                            this._assetsid=x.Id;console.log(x.Id,x.Title);
                            break;
                        case 'TicketNotes':
                            this._ticketnotesid=x.Id;console.log(x.Id,x.Title);
                            break;
                        default:
                            console.log(x.RootFolder.Name);
                            break;
                    }
                });

                const configdata={
                    siteurl: url,
                    ticketid: this._ticketsid,
                    statusid: this._statusid,
                    teamsid: this._teamsid,
                    prioritesid:this._prioritesid,
                    servicegroupid:this._servicegroupid,
                    servicesid:this._servicesid,
                    subcategory:this._subcategory,
                    conversationid:this._conversationid,
                    CIid:this._CIid,
                    assetsid:this._assetsid,
                    ticketnotesid:this._ticketnotesid
                };
                this.setStorageEntity(configdata);
                this.getUsers(null);
                this.getCIsLookUp().then((cidata)=>{
                    this._CIs=cidata;
                });
                return Promise.resolve("success");
            }).catch((ex) => {
                console.log("Error while fetching User Details: ", ex);
                throw ex;
            });
    }

    public getITSMTickets(nexturl?:string,prevtickets?:ITicketItem[],cview?:string):Promise<any>{
        const selectquery:string="$select=ID,Title,SLAPriority/Title,Requester/Title,TicketsStatus/Title,ContentType/name,AssignedPerson/Title,AssignedPerson/EMail,AssignedTeam/Title,Created,TimeToFixModern,Modified,Editor/Title";
        const expandquery:string="$expand=Requester,SLAPriority,TicketsStatus,ContentType,AssignedPerson,AssignedTeam,Editor";
        let querygetAllItems=null;
        if(cview){
            switch(cview){
                case "myview":{
                    const Currentuser:IUserDetails[]=this._lusers.filter(i=>i.Email==this._currentuser.email);
                    const currentuserid:string=Currentuser.length>0?Currentuser[0].ID:"-1";
                    querygetAllItems=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$filter=AssignedPersonStringId eq ${currentuserid} and TicketsStatusId ne 7&$orderby=Id desc`;
                    break;
                }
                case "unassignedview":{
                    querygetAllItems=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$filter= AssignedPersonStringId eq null and AssignedTeamId eq null and TicketsStatusId ne 7&$orderby=Id desc`;
                    break;
                }
                case "openview":{
                    querygetAllItems=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$filter=TicketsStatusId ne 7 and TicketsStatusId ne 12 and TicketsStatusId ne 14&$orderby=Id desc`;
                    break;
                }
                case "allview":{
                    querygetAllItems=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$orderby=Id desc`;
                    break;
                }
                default:{
                    querygetAllItems=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$orderby=Id desc`;
                    break;
                }
            }
        }else{
            querygetAllItems = nexturl?nexturl:`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$orderby=Id desc`;
        }
        
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=verbose"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
               return {"tickets":this.processtickets(data.d.results,prevtickets),
               "nexturl":data.d.__next
               };
            }).catch((ex) => {
                console.log("Error while fetching ITSM tickets: ", ex);
                throw ex;
            });
    }

    //Method for converting the SP data to model data
    private processtickets(items:any[],prevItems:ITicketItem[]):any[]{
        let tickets:ITicketItem[]=prevItems;
        items.forEach((item)=>{
            let asp:string="";
            let pcolor:string="";
            let at:string="";
            if(typeof item.AssignedTeam.Title=="undefined"){
                asp=typeof item.AssignedPerson.Title=="undefined"?"":item.AssignedPerson.Title;
            }else{
                asp=typeof item.AssignedPerson.Title=="undefined"?`${item.AssignedTeam.Title}`:`${item.AssignedTeam.Title}:${item.AssignedPerson.Title}`;
                at=item.AssignedTeam.Title;
            }

            const SPT=typeof item.SLAPriority !="undefined"?item.SLAPriority.Title:"-1";
            switch(SPT){
                case "1":{
                    pcolor="#B21E29";
                    break;
                }
                case "2":{
                    pcolor="#ED7D31";
                    break;
                }
                case "3":{
                    pcolor="#EEC400";
                    break;
                }
                case "4":{
                    pcolor="#7ABC32";
                    break;
                }
                case "5":{
                    pcolor="#456B2B";
                    break;
                }
                default:{
                    pcolor="#fff";
                    break;
                }
            }

            let ticket:ITicketItem={
                key:item.ID,
                ID:item.ID,
                Title:item.Title,
                Priority:SPT,
                Prioritycolor:pcolor,
                Requester:item.Requester.Title,
                Status:item.TicketsStatus.Title,
                ContentType:item.ContentType.Name,
                AssignedTeamPerson:asp,
                AssignedPerson:typeof item.AssignedPerson!="undefined"?item.AssignedPerson.EMail:"",
                Created:moment(item.Created).format("MMM Do YY"),
                RemainingTime:"",
                lastmodified:item.Modified,
                lastmodifiedby:typeof item.Editor!="undefined"?item.Editor.Title:"",
                AssignedTeam:at
            };
            tickets.push(ticket);
        });
        return tickets;
    }

    public getlookupdata():Promise<any>{
        const spBatchCreateOpts:ISPHttpClientBatchCreationOptions={webUrl:this._weburl};
        const spBatch:SPHttpClientBatch=this._spclient.beginBatch(spBatchCreateOpts);
        const teamsurl=`${this._weburl}_api/web/lists(guid'${this._teamsid}')/items?$select=Title,ID`;
        const getteams:Promise<SPHttpClientResponse>=spBatch.get(teamsurl,SPHttpClientBatch.configurations.v1);
        const priorityurl=`${this._weburl}_api/web/lists(guid'${this._prioritesid}')/items`;
        const getpriority:Promise<SPHttpClientResponse>=spBatch.get(priorityurl,SPHttpClientBatch.configurations.v1);
        const statusurl=`${this._weburl}_api/web/lists(guid'${this._statusid}')/items`;
        const getStatus:Promise<SPHttpClientResponse>=spBatch.get(statusurl,SPHttpClientBatch.configurations.v1);
        const conurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/ContentTypes?$select=name,ID`;
        const getTypes:Promise<SPHttpClientResponse>=spBatch.get(conurl,SPHttpClientBatch.configurations.v1);
        return spBatch.execute().then(()=>{
            getpriority.then((response:SPHttpClientResponse)=>{
                response.json().then((pitems)=>{
                    pitems.value.forEach((pitem)=>{
                        let spri:ISLAPriority={
                            ID:pitem.ID,
                            Description:pitem.Description,
                            Title:pitem.Title
                        };
                        this._spris.push(spri);
                    });
                });
            });
            getStatus.then((response:SPHttpClientResponse)=>{
                response.json().then((sitems)=>{
                    sitems.value.forEach((sitem)=>{
                        let stat:Istatus={
                            ID:sitem.ID,
                            Title:sitem.Title
                        };
                        this._stats.push(stat);
                    });
                });
            });
            getTypes.then((response:SPHttpClientResponse)=>{
                response.json().then((titems)=>{
                    titems.value.forEach((titem)=>{
                        let scont:IContype={
                            Name:titem.Name,
                            ID:titem.Id.StringValue
                        };
                        this._sconts.push(scont);
                    });
                });
            });
            getteams.then((response:SPHttpClientResponse)=>{
                response.json().then((steams)=>{
                    steams.value.forEach((steam)=>{
                        let stea:ITeam={
                            ID:steam.ID,
                            Title:steam.Title
                        };
                        this._steams.push(stea);
                    });
                });
            });
            return Promise.all([this._sconts,this._spris,this._stats,this._steams]);
        });
    }

    public getSearchResults(filters:any):Promise<any>{
        let fils:string="";
        let pfils:string="";
        let sfils:string="";
        let ctfils:string="";
        let tfils:string="";
        let ifils:string="";
        if(typeof filters.Priority!="undefined"){
            filters.Priority.forEach((filter,key)=>{
                if(pfils.length<=1){
                    pfils=`<Eq><FieldRef Name="SLAPriority"/><Value Type="Lookup">${filter}</Value></Eq>`;
                }else{
                    pfils=`<Or>${pfils}<Eq><FieldRef Name="SLAPriority"/><Value Type="Lookup">${filter}</Value></Eq></Or>`;
                }
            });
        }

        if(pfils.length>0){
            fils=pfils;
        }

        if(typeof filters.Status!="undefined"){
            filters.Status.forEach((filter,key)=>{
                if(sfils.length<=1){
                    sfils=`<Eq><FieldRef Name="TicketsStatus"/><Value Type="Lookup">${filter}</Value></Eq>`;
                }else{
                    sfils=`<Or>${sfils}<Eq><FieldRef Name="TicketsStatus"/><Value Type="Lookup">${filter}</Value></Eq></Or>`;
                }
            });
        }

        if(sfils.length>0 && fils.length>0){
            fils=`<And>${fils}${sfils}</And>`;
        }else if(sfils.length>0){
            fils=sfils;
        }

        if(typeof filters.ContentType!="undefined"){
            filters.ContentType.forEach((filter,key)=>{
                if(ctfils.length<=1){
                    ctfils=`<BeginsWith><FieldRef Name="ContentTypeId"/><Value Type="ContentTypeId">${filter}</Value></BeginsWith>`;
                }else{
                    ctfils=`<Or>${ctfils}<BeginsWith><FieldRef Name="ContentTypeId"/><Value Type="ContentTypeId">${filter}</Value></BeginsWith></Or>`;
                }
            });
        }

        if(ctfils.length>0 && fils.length>0){
            fils=`<And>${fils}${ctfils}</And>`;
        }else if(ctfils.length>0){
            fils=ctfils;
        }

        if(typeof filters.Title!="undefined"){
            filters.Title.forEach((filter,key)=>{
                if(tfils.length<=1){
                    tfils=`<Contains><FieldRef Name="Title"/><Value Type="Text">${filter}</Value></Contains>`;
                }else{
                    tfils=`<Or>${tfils}<Contains><FieldRef Name="Title"/><Value Type="Text">${filter}</Value></Contains></Or>`;
                }
            });
        }

        if(tfils.length>0 && fils.length>0){
            fils=`<And>${fils}${tfils}</And>`;
        }else if(tfils.length>0){
            fils=tfils;
        }

        if(typeof filters.ID!="undefined"){
            filters.ID.forEach((filter,key)=>{
                if(ifils.length<=1){
                    ifils=`<Eq><FieldRef Name="ID"/><Value Type="Number">${filter}</Value></Eq>`;
                }else{
                    ifils=`<Or>${ifils}<Eq><FieldRef Name="ID"/><Value Type="Number">${filter}</Value></Eq></Or>`;
                }
            });
        }

        if(ifils.length>0 && fils.length>0){
            fils=`<And>${fils}${ifils}</And>`;
        }else if(ifils.length>0){
            fils=ifils;
        }

        const itemurl:string=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/GetItems`;
        const options: ISPHttpClientOptions = {
            headers: {
                'odata-version':'3.0',
                'accept':'application/json;odata=nometadata'
                    },
            body: `{'query': {
                '__metadata': {'type': 'SP.CamlQuery'},
                'ViewXml': '<View><Query>
                <Where>${fils}</Where>
                <OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy></Query>
                <ViewFields>
                <FieldRef Name="ID" />
                <FieldRef Name="Title" />
                <FieldRef Name="SLAPriority" />
                <FieldRef Name="Requester" />
                <FieldRef Name="TicketsStatus" />
                <FieldRef Name="ContentTypeId" />
                <FieldRef Name="AssignedPerson" />
                <FieldRef Name="AssignedTeam" />
                <FieldRef Name="Created" />
                <FieldRef Name="TimeToFixModern" />
                </ViewFields>
                <RowLimit>100</RowLimit></View>'
            }}`
        };

        return this._spclient.post(itemurl,SPHttpClient.configurations.v1,options).then((response:SPHttpClientResponse)=>{
            if (response.status >= 200 && response.status < 300) {
                return response.json();
            }
            else { return Promise.reject(new Error(JSON.stringify(response))); }
        }).then((data: any) => {
            let tickets:ITicketItem[]=[];
            data.value.forEach((item)=>{
            let pcolor:string="";
            const lpri:ISLAPriority[]=this._spris.filter(i => i.ID == item.SLAPriorityId);
            const lstat:Istatus[]=this._stats.filter(i=>i.ID==item.TicketsStatusId);
            const lct:IContype[]=this._sconts.filter(i=>i.ID==item.ContentTypeId);
            
            switch(lpri.length>0?lpri[0].Title:"-1"){
                case "1":{
                    pcolor="#B21E29";
                    break;
                }
                case "2":{
                    pcolor="#ED7D31";
                    break;
                }
                case "3":{
                    pcolor="#EEC400";
                    break;
                }
                case "4":{
                    pcolor="#7ABC32";
                    break;
                }
                case "5":{
                    pcolor="#456B2B";
                    break;
                }
                default:{
                    pcolor="#fff";
                    break;
                }
            }

            let asp:string="";
            const ltm:ITeam[]=this._steams.filter(i=>i.ID==item.AssignedTeamId);
            const ltmval=ltm.length>0?ltm[0].Title:"";
                let rusers=this._lusers.filter(i=>i.ID==item.RequesterStringId);
                let ausers=this._lusers.filter(i=>i.ID==item.AssignedPersonStringId);
                if(ltmval.length>0){
                    asp=ausers.length>0?`${ltmval}:${ausers[0].Title}`:ltmval;
                }else{
                    asp=ausers.length>0?ausers[0].Title:"";
                }
                let ticket:ITicketItem={
                    key:item.ID,
                    ID:item.ID,
                    Title:item.Title,
                    Priority:lpri.length>0?lpri[0].Title:"",
                    Prioritycolor:pcolor,
                    Requester:rusers.length>0?rusers[0].Title:"",
                    Status:lstat.length>0?lstat[0].Title:"",
                    ContentType:lct.length>0?lct[0].Name:"",
                    AssignedTeamPerson:asp,
                    AssignedPerson:ausers.length>0?ausers[0].Email:"",
                    Created:moment(item.Created).format("MMM Do YY"),
                    RemainingTime:"",
                    AssignedTeam:ltmval
                };
                tickets.push(ticket);
        });

        return tickets;
        }).catch((ex)=>{
            console.log(ex);
        });
    }

    private getUsers(url:string):Promise<void>{
        const querygetAllUsers = url?url:`${this._weburl}_api/web/SiteUserInfoList/items?&$select=Id,Title,Name,EMail,Picture`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=verbose"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllUsers, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                data.d.results.forEach((user)=>{
                    //console.log(user.Picture?user.Picture.Url:"");
                    let luser:IUserDetails={
                        ID:user.Id,
                        Title:user.Title,
                        Name:user.Name,
                        Email:user.EMail,
                        pictureurl:user.Picture?user.Picture.Url:""
                    };
                    this._lusers.push(luser);
                });
                if(data.d.__next){
                    this.getUsers(data.d.__next);
                }
            }).catch((ex) => {
                console.log("Error while fetching User Details: ", ex);
                throw ex;
            });
    }

    public getMyTicketsCount():Promise<any>{
        const Currentuser:IUserDetails[]=this._lusers.filter(i=>i.Email==this._currentuser.email);
        const currentuserid:string=Currentuser.length>0?Currentuser[0].ID:"-1";
        const myticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID&$filter=AssignedPersonStringId eq ${currentuserid} and TicketsStatusId ne 7`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(myticketsurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    public getUnassignedTicketsCount():Promise<any>{
        const unassignedticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID&$filter= AssignedPersonStringId eq null and AssignedTeamId eq null and TicketsStatusId ne 7`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(unassignedticketsurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    //https://codewithrohit.wordpress.com/2017/06/01/sharepoint-rest-api/
    public getopenTicketsCount():Promise<any>{
        const openticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID,TicketsStatusId&$filter=TicketsStatusId ne 7 and TicketsStatusId ne 12 and TicketsStatusId ne 14&$top=5000`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    public getallticketscount():Promise<any>{
        const allticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')?$select=Itemcount`;
        return this._spclient.get(allticketsurl, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.ItemCount;
            }).catch((ex) => {
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    public getlookupdatanew():Promise<any>{
        const spBatchCreateOpts:ISPHttpClientBatchCreationOptions={webUrl:this._weburl};
        const spBatch:SPHttpClientBatch=this._spclient.beginBatch(spBatchCreateOpts);
        const sgurl=`${this._weburl}_api/web/lists(guid'${this._servicegroupid}')/items?$select=Title,ID`;
        const getServieGroups:Promise<SPHttpClientResponse>=spBatch.get(sgurl,SPHttpClientBatch.configurations.v1);
        const srurl=`${this._weburl}_api/web/lists(guid'${this._servicesid}')/items?$select=Title,ID,ServiceGroup/Title&$expand=ServiceGroup`;
        const getServies:Promise<SPHttpClientResponse>=spBatch.get(srurl,SPHttpClientBatch.configurations.v1);
        const scgurl=`${this._weburl}_api/web/lists(guid'${this._subcategory}')/items?$select=ID,Title,RelatedBusinessService/Title&$expand=RelatedBusinessService`;
        const getSubcategory:Promise<SPHttpClientResponse>=spBatch.get(scgurl,SPHttpClientBatch.configurations.v1);
        return spBatch.execute().then(()=>{
            let sgroups:IServiceGroup[]=[];
            let services:IService[]=[];
            let scategories:IServiceCategory[]=[];
            let rdatas:any[]=[];
             return getServieGroups.then((response:SPHttpClientResponse)=>{
                return response.json().then((sgitems)=>{
                    sgitems.value.forEach((sgitem)=>{
                        let sgroup:IServiceGroup={
                            ID:sgitem.ID,
                            Title:sgitem.Title
                        };
                        sgroups.push(sgroup);
                    });

                    return getServies.then((sresponse:SPHttpClientResponse)=>{
                        return sresponse.json().then((sitems)=>{
                            sitems.value.forEach((sitem)=>{
                                let service:IService={
                                    ID:sitem.ID,
                                    Title:sitem.Title,
                                    ServiceGroup:sitem.ServiceGroup.Title
                                };
                                services.push(service);
                            });

                            return getSubcategory.then((scresponse:SPHttpClientResponse)=>{
                                return scresponse.json().then((scitems)=>{
                                    scitems.value.forEach((scitem)=>{
                                        let scategorie:IServiceCategory={
                                            ID:scitem.ID,
                                            Title:scitem.Title,
                                            Service:scitem.RelatedBusinessService.Title
                                        };
                                        scategories.push(scategorie);
                                    });
                                
                                sgroups.forEach((sgrp)=>{
                                    let x:IService[]=services.filter(i => i.ServiceGroup == sgrp.Title);
                                    let xitems:any[]=[];
                                    x.forEach((ss)=>{
                                        let y:IServiceCategory[]=scategories.filter(i=>i.Service==ss.Title);
                                        let yitems:any[]=[];
                                        y.forEach((sc)=>{
                                            let yitem={
                                                value:sc.ID,
                                                label:sc.Title
                                            };
                                            yitems.push(yitem);
                                        });
                                        let xitem={
                                            value:ss.ID,
                                            label:ss.Title,
                                            children:yitems
                                        };
                                        xitems.push(xitem);
                                    });

                                    let rdata={
                                        value:sgrp.ID,
                                        label:sgrp.Title,
                                        children:xitems
                                    };
                                    rdatas.push(rdata);
                                });

                                return rdatas;
                            });
                            });
                        });     
                    });  
                });
            });
        });
    }

    public addITSMTicket(itsmticket:any):Promise<any>{
        const addcaseurl:string=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            body:JSON.stringify(itsmticket)
        };

        return this._spclient.post(addcaseurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public uploadTicketAttachment(file:any,ticketid:String):Promise<any>{
        const uploadurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})/AttachmentFiles/add(FileName='${file.name}')`;
        let spOpts:ISPHttpClientOptions={
            headers:{
                "Accept":"application/json",
                "Content-Type":"application/json"
            },
            body:file
        };

        return this._spclient.post(uploadurl,SPHttpClient.configurations.v1,spOpts).then((response:SPHttpClientResponse)=>{
            response.json().then((responseJSON:JSON)=>{
                return responseJSON;
            });
        });
    }

    public getTicketAttachment(ticketid:string):Promise<any>{
        const webUrl="https://cloudmission.sharepoint.com";
        const ticketattachmenturl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})/AttachmentFiles?$select=FileName,ServerRelativeUrl`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };

        return this._spclient.get(ticketattachmenturl,SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            }).then((data: any) => {
                let ticketattachments:any[]=[];
                data.value.forEach(attach => {
                   const ticketattach:any={
                       filename:attach.FileName,
                       attachurl:`${webUrl}${attach.ServerRelativeUrl}`
                   };
                   ticketattachments.push(ticketattach);
                });
                return ticketattachments;
            }).catch((ex) => {
                console.log("Error while fetching Ticket attachments: ", ex);
                throw ex;
            });
    }

    public updateTicketStatus(ticketid:string,statusid:string):Promise<any>{
        const updateurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify({
                    'TicketsStatusId': statusid
                  });
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json",
                        "Content-Type":"application/json",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
            
          }).catch((ex) => {
                console.log("Error while updating status: ", ex);
                throw ex;
            });
    }

    public updateTicketAssign(ticketid:string,updateobj:any):Promise<any>{
        const updateurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify(updateobj);
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json",
                        "Content-Type":"application/json",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
            
          }).catch((ex) => {
                console.log("Error while updating status: ", ex);
                throw ex;
            });
    }

    //Method for getting the Ticket communications. Shown as conversations.
    public getTicketNotes(ticketid:string):Promise<any>{
        const ticketnotesurl=`${this._weburl}_api/web/lists(guid'${this._conversationid}')/items?$filter=TicketIDId eq ${ticketid}&$select=ID,AuthorId,Communications,CommunicationInitiatorId,Created&$orderby=Id desc`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(ticketnotesurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let ticketnotes:any[]=[];
                data.value.forEach(note => {
                   const users=this._lusers.filter(i=>i.ID==note.CommunicationInitiatorId);
                   const user=users.length>0?users[0]:null;
                   const ticketnote:any={
                       author:user.Title,
                       avatar:`${this._weburl}_layouts/15/userphoto.aspx?size=S&username=${user.Email}`,
                       content:note.Communications,
                       datetime:note.Created
                   };
                   ticketnotes.push(ticketnote);
                });
                return ticketnotes;
            }).catch((ex) => {
                console.log("Error while fetching Ticket Notes: ", ex);
                throw ex;
            });
    }

    public addTicketNotes(ticketnote:any):Promise<any>{
        const addnotesurl:string=`${this._weburl}_api/web/lists(guid'${this._conversationid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            headers:{
                "Accept":"application/json;odata=verbose",
                "Content-Type":"application/json;odata=verbose",
                "odata-version": ""
            },
            body:JSON.stringify(ticketnote)
        };

        return this._spclient.post(addnotesurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public getTicketRequesterAndSenderEmails(ticketid:string):Promise<any>{
        const ticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$filter=ID eq ${ticketid}&$select=Requester/ID,Sender&$expand=Requester`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(ticketsurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let requesterUser = (data.value.length > 0 && data.value[0].Requester != null && data.value[0].Requester != undefined) ? this._lusers.filter(i=>i.ID==data.value[0].Requester.ID) : null;
                let requesterEmail = (requesterUser != null && requesterUser.length > 0) ? requesterUser[0].Email : "";
                let senderEmail = (data.value.length > 0 && data.value[0].Sender != null && data.value[0].Sender != undefined) ? data.value[0].Sender : "";
                return { "Requester": requesterEmail, "Sender": senderEmail };
            }).catch((ex) => {
                console.log("Error while fetching current ticket Requester: ", ex);
                throw ex;
            });
    }

    public getTicketEmails(ticketid:string):Promise<any>{
        const emailsurl=`${this._weburl}_api/web/lists(guid'${this._emailsid}')/items?$filter=RelatedItem eq ${ticketid} and RelatedList eq 'Tickets'&$select=ID,Title,Email,Received,Created,Message,Cc,Read&$orderby=Id desc`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(emailsurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let ticketEmails:any[]=[];
                data.value.forEach(email => {
                    email.Created = moment(email.Created).format("MMM Do YY hh:mm");
                    ticketEmails.push(email);
                });
                return ticketEmails;
            }).catch((ex) => {
                console.log("Error while fetching Ticket Emails: ", ex);
                throw ex;
            });
    }

    public addTicketEmail(ticketEmail:any):Promise<any>{
        const emailsUrl:string=`${this._weburl}_api/web/lists(guid'${this._emailsid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            body:JSON.stringify(ticketEmail)
        };

        return this._spclient.post(emailsUrl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public getTicketInternalNotes(ticketid:string):Promise<any>{
        const ticketnotesurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})/versions/?$select=ID,Notes,VersionLabel,created,Author&$Orderby=VersionLabel desc`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(ticketnotesurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let ticketnotes:any[]=[];
                const aitems=data.value;
                for (let i = 0; i < aitems.length; i++) {
                    if (aitems[i].Notes != null) {
                        const ticketnote: any = {
                            author: aitems[i].Author.LookupValue,
                            avatar: `${this._weburl}_layouts/15/userphoto.aspx?size=S&username=${aitems[i].Author.Email}`,
                            content: aitems[i].Notes,
                            datetime: aitems[i].Created
                        };
                        ticketnotes.push(ticketnote);
                    }
                }
                return ticketnotes;
            }).catch((ex) => {
                console.log("Error while fetching Ticket Notes: ", ex);
                throw ex;
            });
    }

    public getTicketDetails(ticketid:String):Promise<any>{
        const selectquery:string="$select=Description,RequestSummary,Urgency,Impact,ServiceGroups/Title,ServiceGroups/ID,RelatedServices/Title,RelatedServices/ID,RelatedCategories/Title,RelatedCategories/ID,RelatedCIs/Title,RelatedCIs/ID,RelatedAssets/ID,NotificationSummary,OrderDetails";
        const expandquery:string="$expand=ServiceGroups,RelatedServices,RelatedCategories,RelatedCIs,RelatedAssets";
        //const filterquery:string=`$filter=ID eq ${ticketid}`
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})?${selectquery}&${expandquery}`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=verbose"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                console.log("ticket details ",data.d);
               return data.d;
               
            }).catch((ex) => {
                console.log("Error while fetching ITSM tickets: ", ex);
                throw ex;
            });
    }

    public updateTicketDetails(itsmticket:any,ticketid:string):Promise<any>{
        const updateurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify(itsmticket);
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json;odata=verbose",
                        "Content-Type":"application/json;odata=verbose",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
          }).catch((ex) => {
                console.log("Error while updating ticket details: ", ex);
                throw ex;
            });
    }

    public PostToTeams(ticket:any):Promise<any>{
        const flowurl="https://prod-118.westeurope.logic.azure.com:443/workflows/e67a8cb8aefc45159ec946e8d4a9b3bf/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=deziuFc9wZlJkaItf7JNLN35glXJ4HLMdPFZnmNpIlc";
        const requester=this._lusers.filter(i=>i.ID==ticket.RequesterId);
        const tname=this._steams.filter(i=>i.ID==ticket.AssignedTeamId);
        let ap:IUserDetails[]=[];
        if(typeof ticket.AssignedPersonId !="undefined"){
            ap=this._lusers.filter(i=>i.ID==ticket.AssignedPersonId);
          }
        const body:string=JSON.stringify({
            "teamid":this._teamscontext.groupId,
            "channelid":this._teamscontext.channelId,
            "TicketID":"",
            "TicketTitle":ticket.Title,
            "RequestedBy":requester.length>0?requester[0].Title:"",
            "TicketDescription":ticket.Description,
            "AssignedTeam":tname.length>0?tname[0].Title:"",
            "AssignedPerson":ap.length>0?ap[0].Title:"",
            "Urgency":ticket.Urgency,
            "Impact":ticket.Impact,
            "pictureurl":requester.length>0?`${this._weburl}_layouts/15/userphoto.aspx?size=S&username=${requester[0].Email}`:""
        });    
        const requestHeaders: Headers = new Headers();
        requestHeaders.append('Content-type', 'application/json');

        const httpClientOptions: IHttpClientOptions = {
            body: body,
            headers: requestHeaders
          };

          return this._httpclient.post(flowurl,HttpClient.configurations.v1,httpClientOptions).then((response: HttpClientResponse) => {
            if (response.status >= 200 && response.status < 300) {
                return response.status;
            }
            else { return Promise.reject(new Error(JSON.stringify(response))); }
        });
    }

    public getCIsLookUp():Promise<any>{
        const ciurl=`${this._weburl}_api/web/lists(guid'${this._CIid}')/items?$select=ID,Title`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(ciurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let cis:ICI[]=[];
                data.value.forEach(lci => {
                   const ci:ICI={
                    Title:lci.Title,
                    ID:lci.ID
                   };
                   cis.push(ci);
                });
                return cis;
            }).catch((ex) => {
                console.log("Error while fetching CIs lookupdata: ", ex);
                throw ex;
            });
    }

    public getuserAssets(ticket:ITicketItem):Promise<any>{
        const requestor=this._lusers.filter(i => i.Title == ticket.Requester);
        const userid=requestor.length>0?requestor[0].ID:"-1";
        const selectquery:string="$select=ID,Title,Model/Title,LifeCycleStage/Title,ContentType/name";
        const expandquery:string="$expand=Model,LifeCycleStage,ContentType";
        const filterquery:string=`$filter= EndUserId eq ${userid} and LifeCycleStageId ne 4`;
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._assetsid}')/items?${selectquery}&${expandquery}&${filterquery}`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let assets:IAsset[]=[];
                data.value.forEach(lasset => {
                   const asset:IAsset={
                    Title:lasset.Title,
                    Model:lasset.Model.Title,
                    State:lasset.LifeCycleStage.Title,
                    Type:lasset.ContentType.Name,
                    key:lasset.ID,
                    ID:lasset.ID
                   };
                   assets.push(asset);
                });
                console.log(assets);
                return assets;
            }).catch((ex) => {
                console.log("Error while fetching assets: ", ex);
                throw ex;
            });
    }

    public updateTicketRelatedAssets(assetdetails:any,ticketid:string):Promise<any>{
        const updateurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items(${ticketid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify(assetdetails);
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json;odata=verbose",
                        "Content-Type":"application/json;odata=verbose",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
          }).catch((ex) => {
                console.log("Error while updating ticket details: ", ex);
                throw ex;
            });
    }

    public AddTicketInternalNotes(ticketnote:any):Promise<any>{
        const addnotesurl:string=`${this._weburl}_api/web/lists(guid'${this._ticketnotesid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            headers:{
                "Accept":"application/json;odata=verbose",
                "Content-Type":"application/json;odata=verbose",
                "odata-version": ""
            },
            body:JSON.stringify(ticketnote)
        };

        return this._spclient.post(addnotesurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public getSubTasks(ticketid:string):Promise<any>{
        const tskurl=`${this._weburl}_api/web/lists(guid'${this._tickettasksid}')/items?$select=ID,Title,AssignedPerson/Title,AssignedPerson/EMail,AssignedTeam/Title,AssignedTeam/Id,StartDate,DueDate,PercentComplete,TaskStatus,Description&$expand=AssignedPerson,AssignedTeam&$filter=RelatedTicketsId eq ${ticketid}`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(tskurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let tsks:ITask[]=[];
                data.value.forEach(ltsk => {
                   const tsk:ITask={
                    key:ltsk.ID,   
                    Title:ltsk.Title,
                    StartDate:ltsk.StartDate,
                    DueDate:ltsk.DueDate,
                    PercentComplete:ltsk.PercentComplete,
                    TaskStatus:ltsk.TaskStatus,
                    Description:ltsk.Description,
                    AssignedTeam:typeof ltsk.AssignedTeam!= 'undefined'?{Title:ltsk.AssignedTeam.Title,ID:ltsk.AssignedTeam.Id}:null,
                    AssignedPerson:typeof ltsk.AssignedPerson!= 'undefined'?{Title:ltsk.AssignedPerson.Title,ID:ltsk.AssignedPerson.EMail}:null
                   };
                   tsks.push(tsk);
                });
                return tsks;
            }).catch((ex) => {
                console.log("Error while fetching Ticket SubTasks data: ", ex);
                throw ex;
            });
    }

    public updateITSMSubTask(itsmtask:any,tskid:string):Promise<any>{
        const updateurl=`${this._weburl}_api/web/lists(guid'${this._tickettasksid}')/items(${tskid})`;
        const getetagurl=`${this._weburl}_api/web/lists(guid'${this._tickettasksid}')/items(${tskid})?$select=Id`;
        let etag: string = undefined;
        return this._spclient.get(getetagurl,SPHttpClient.configurations.v1,{
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response:SPHttpClientResponse)=>{
            etag=response.headers.get("ETag");
            return response.json().then((rdata)=>{
                const body:string=JSON.stringify(itsmtask);
                 const data:ISPHttpClientBatchOptions={
                    headers:{
                        "Accept":"application/json;odata=verbose",
                        "Content-Type":"application/json;odata=verbose",
                        "odata-version": "",
                        "IF-MATCH": etag,
                        "X-HTTP-Method": "MERGE"
                    },
                    body:body
                 };
                 return this._spclient.post(updateurl,SPHttpClient.configurations.v1,data).then((postresponse:SPHttpClientResponse)=>{
                    return postresponse;
                 });
            });
          }).catch((ex) => {
                console.log("Error while updating task details: ", ex);
                throw ex;
            });
    }

    public addITSMSubTask(itsmtask:any):Promise<any>{
        const addtaskurl:string=`${this._weburl}_api/web/lists(guid'${this._tickettasksid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            body:JSON.stringify(itsmtask)
        };

        return this._spclient.post(addtaskurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public getTaskCatalog():Promise<any>{
        const catalogurl=`${this._weburl}_api/web/lists(guid'${this._taskCatalog}')/items?$select=ID,Title,AssignedPerson/Title,AssignedPerson/EMail,AssignedTeam/Title,AssignedTeam/Id,Description&$expand=AssignedPerson,AssignedTeam`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(catalogurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let tsks:ITask[]=[];
                data.value.forEach(ltsk => {
                   const tsk:ITask={
                    key:ltsk.ID,   
                    Title:ltsk.Title,
                    Description:ltsk.Description,
                    AssignedTeam:typeof ltsk.AssignedTeam!= 'undefined'?{Title:ltsk.AssignedTeam.Title,ID:ltsk.AssignedTeam.Id}:null,
                    AssignedPerson:typeof ltsk.AssignedPerson!= 'undefined'?{Title:ltsk.AssignedPerson.Title,ID:ltsk.AssignedPerson.EMail}:null
                   };
                   tsks.push(tsk);
                });
                return tsks;
            }).catch((ex) => {
                console.log("Error while fetching Ticket SubTasks data: ", ex);
                throw ex;
            });
    }

    public getLanguages():Promise<any>{
        const langurl=`${this._weburl}_api/web/lists(guid'${this._languageid}')/items?$select=ID,Title,DefaultLanguage`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(langurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let langs:any[]=[];
                data.value.forEach(lci => {
                   const lang:any={
                    Title:lci.Title,
                    ID:lci.ID,
                    DefaultLanguage:lci.DefaultLanguage
                   };
                   langs.push(lang);
                });
                return langs;
            }).catch((ex) => {
                console.log("Error while fetching language data: ", ex);
                throw ex;
            });
    }

    public addKBArticle(KBObj):Promise<any>{
        const addtaskurl:string=`${this._weburl}_api/web/lists(guid'${this._KBArtcileid}')/items`;
        const httpclientoptions:ISPHttpClientOptions={
            headers:{
                "Accept":"application/json;odata=verbose",
                "Content-Type":"application/json;odata=verbose",
                "odata-version": ""
            },
            body:JSON.stringify(KBObj)
        };

        return this._spclient.post(addtaskurl, SPHttpClient.configurations.v1, httpclientoptions)
            .then((response: SPHttpClientResponse) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.status;
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            });
    }

    public getStandardMessages():Promise<any>{
        const selectquery:string="$select=Title,Message,ID,RelatedService/Title,RelatedService/ID,RelatedCategory/Title,RelatedCategory/ID";
        const expandquery:string="$expand=RelatedService,RelatedCategory";
        //const filterquery:string=`$filter=ID eq ${ticketid}`
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._standMessage}')/items?${selectquery}&${expandquery}`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let msgs:any[]=[];
                data.value.forEach(lci => {
                   const msg:any={
                    Title:lci.Title,
                    ID:lci.ID,
                    Message:lci.Message,
                    ServiceId:typeof lci.RelatedService !="undefined"?lci.RelatedService.ID:null,
                    CategoryId:typeof lci.RelatedCategory !="undefined"?lci.RelatedCategory.ID:null
                   };
                   msgs.push(msg);
                });
                return msgs;
            }).catch((ex) => {
                console.log("Error while fetching standard messages data: ", ex);
                throw ex;
            });
    }

    public getKBArtciles():Promise<any>{
        const langurl=`${this._weburl}_api/web/lists(guid'${this._KBArtcileid}')/items?$select=ID,Title`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(langurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let KBs:any[]=[];
                data.value.forEach(lci => {
                   const KB:any={
                    Title:lci.Title,
                    ID:lci.ID
                   };
                   KBs.push(KB);
                });
                return KBs;
            }).catch((ex) => {
                console.log("Error while fetching KBArticles data: ", ex);
                throw ex;
            });
    }

    public getsitedocuments():Promise<any>{
        const docurl=`${this._weburl}_api/web/lists(guid'${this._doclibraryid}')/items?$select=ID,FileLeafRef,PublishForEndUsers&$filter=PublishForEndUsers eq 1`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(docurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                let docs:any[]=[];
                data.value.forEach(lci => {
                   const doc:any={
                    Title:lci.FileLeafRef,
                    ID:lci.ID
                   };
                   docs.push(doc);
                });
                return docs;
            }).catch((ex) => {
                console.log("Error while fetching KBArticles data: ", ex);
                throw ex;
            });
    }

    public getAssetsAvailability(atitle:string):Promise<any>{
        const aurl=`${this._weburl}_api/web/lists(guid'${this._assetsid}')/items?$select=ID,Title,LifeCycleStage/Title&$filter=Title eq '${atitle}' and LifeCycleStage/Title eq 'In stock'&$expand=LifeCycleStage`;
        const options:ISPHttpClientOptions={
            headers:{
                "odata-version":"3.0",
                "accept":"application/json;odata=nometadata"
            },
            method:"GET"
        };
        return this._spclient.get(aurl, SPHttpClient.configurations.v1,options).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                console.log("Error while fetching Assets details: ", ex);
                throw ex;
            });
    }

}
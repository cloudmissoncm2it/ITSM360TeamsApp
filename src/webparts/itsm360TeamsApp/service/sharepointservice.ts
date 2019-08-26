import { SPHttpClient,  SPHttpClientConfiguration, SPHttpClientResponse, ODataVersion, ISPHttpClientConfiguration, ISPHttpClientOptions, ISPHttpClientBatchOptions, SPHttpClientBatch, ISPHttpClientBatchCreationOptions } from "@microsoft/sp-http";
import { ITicketItem } from "../model/ITicketItem";
import * as moment from 'moment';
import { ISLAPriority } from "../model/ISLAPriority";
import { Istatus } from "../model/Istatus";
import { IContype } from "../model/IContype";
import { ITeam } from "../model/ITeam";
import { IUserDetails } from "../model/IUserDetails";
import { SPUser } from "@microsoft/sp-page-context";
import { IServiceGroup } from "../model/IServiceGroup";
import { IService } from "../model/IService";
import { IServiceCategory } from "../model/IServiceCategory";
import { Children } from "react";

export class sharepointservice{
    private _spclient:SPHttpClient;
    private _weburl="https://cloudmission.sharepoint.com/sites/ITSM360Trial/";
    private _ticketsid="ae3bf971-67ad-407e-870a-71a5f6bb27f8";
    private _teamsid="023d0962-ec23-4596-a212-af1afd6781dc";
    private _prioritesid="4b32c8d6-f2b0-43ba-a24b-76fe4535c328";
    private _statusid="69c202e2-f6a7-4ce0-96b6-67d5527f037c";
    private _servicegroupid="c3619f14-b00c-46d0-a3fa-9d373a9bd60e";
    private _servicesid="ea98ea2b-5179-4c18-982b-d1142ca3550f";
    private _subcategory="5cd9db0b-3549-41d9-adb9-a1f28c94a6a2";
    private _spris:ISLAPriority[]=[];
    private _stats:Istatus[]=[];
    private _sconts:IContype[]=[];
    private _steams:ITeam[]=[];
    private _lusers:IUserDetails[]=[];
    public _currentuser:SPUser;

    constructor(spclient:SPHttpClient,user:SPUser){
        this._spclient=spclient;
        this._currentuser=user;
        console.log(this._currentuser.email);
        console.log(this._currentuser.loginName);
        this.getUsers();
    }

    public getITSMTickets():Promise<any>{
        const selectquery:string="$select=ID,Title,SLAPriority/Title,Requester/Title,TicketsStatus/Title,ContentType/name,AssignedPerson/Title,AssignedTeam/Title,Created,TimeToFixModern";
        const expandquery:string="$expand=Requester,SLAPriority,TicketsStatus,ContentType,AssignedPerson,AssignedTeam";
        const querygetAllItems = `${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?${selectquery}&${expandquery}&$orderby=Id desc`;
        return this._spclient.get(querygetAllItems, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
               return this.processtickets(data.value);
            }).catch((ex) => {
                debugger;
                console.log("Error while fetching ITSM tickets: ", ex);
                throw ex;
            });
    }

    //Method for converting the SP data to model data
    private processtickets(items:any[]):any[]{
        let tickets:ITicketItem[]=[];
        items.forEach((item)=>{
            let asp:string="";
            let pcolor:string="";
            if(typeof item.AssignedTeam=="undefined"){
                asp=typeof item.AssignedPerson=="undefined"?"":item.AssignedPerson.Title;
            }else{
                asp=typeof item.AssignedPerson=="undefined"?`${item.AssignedTeam.Title}`:`${item.AssignedTeam.Title}:${item.AssignedPerson.Title}`;
            }
            switch(item.SLAPriority.Title){
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
                Priority:item.SLAPriority.Title,
                Prioritycolor:pcolor,
                Requester:item.Requester.Title,
                Status:item.TicketsStatus.Title,
                ContentType:item.ContentType.Name,
                AssignedTeamPerson:asp,
                Created:moment(item.Created).format("MMM Do YY"),
                RemainingTime:""
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

        const itemurl:string=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/GetItems`;
        const options: ISPHttpClientOptions = {
            headers: {'odata-version':'3.0'},
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
                    ID:item.ID,
                    Title:item.Title,
                    Priority:lpri.length>0?lpri[0].Title:"",
                    Prioritycolor:pcolor,
                    Requester:rusers.length>0?rusers[0].Title:"",
                    Status:lstat.length>0?lstat[0].Title:"",
                    ContentType:lct.length>0?lct[0].Name:"",
                    AssignedTeamPerson:asp,
                    Created:moment(item.Created).format("MMM Do YY"),
                    RemainingTime:""
                };
                tickets.push(ticket);
        });

        return tickets;
        }).catch((ex)=>{
            console.log(ex);
        });
    }

    private getUsers():Promise<void>{
        const querygetAllUsers = `${this._weburl}_api/web/SiteUserInfoList/items?&$select=Id,Title,Name,EMail`;
        return this._spclient.get(querygetAllUsers, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                data.value.forEach((user)=>{
                    let luser:IUserDetails={
                        ID:user.Id,
                        Title:user.Title,
                        Name:user.Name,
                        Email:user.EMail
                    };
                    this._lusers.push(luser);
                });
            }).catch((ex) => {
                debugger;
                console.log("Error while fetching User Details: ", ex);
                throw ex;
            });
    }

    public getMyTicketsCount():Promise<any>{
        const Currentuser:IUserDetails[]=this._lusers.filter(i=>i.Email==this._currentuser.email);
        const currentuserid:string=Currentuser.length>0?Currentuser[0].ID:"-1";
        const myticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID&$filter=AssignedPersonStringId eq ${currentuserid} and TicketsStatusId ne 7`;
        return this._spclient.get(myticketsurl, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                debugger;
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    public getUnassignedTicketsCount():Promise<any>{
        const unassignedticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID&$filter= AssignedPersonStringId eq null and AssignedTeamId eq null and TicketsStatusId ne 7`;
        return this._spclient.get(unassignedticketsurl, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                debugger;
                console.log("Error while fetching My tickets count: ", ex);
                throw ex;
            });
    }

    //https://codewithrohit.wordpress.com/2017/06/01/sharepoint-rest-api/
    public getopenTicketsCount():Promise<any>{
        const openticketsurl=`${this._weburl}_api/web/lists(guid'${this._ticketsid}')/items?$select=ID,TicketsStatusId&$filter=TicketsStatusId ne 7 and TicketsStatusId ne 12 and TicketsStatusId ne 14&$top=5000`;
        return this._spclient.get(openticketsurl, SPHttpClient.configurations.v1).then(
            (response: any) => {
                if (response.status >= 200 && response.status < 300) {
                    return response.json();
                }
                else { return Promise.reject(new Error(JSON.stringify(response))); }
            })
            .then((data: any) => {
                return data.value.length;
            }).catch((ex) => {
                debugger;
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
                debugger;
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
}
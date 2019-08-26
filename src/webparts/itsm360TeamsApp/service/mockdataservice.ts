import { ITicketItem } from "../model/ITicketItem";
import * as moment from 'moment';

export class mockdataservice{

    public getMockData():Promise<any>{
        const datasource:ITicketItem[]=[
            {
                ID:"18",
                Title:"How can we improve security?",
                Priority:"3",
                Prioritycolor:"#EEC400",
                Requester:"Kristine Bjorkman",
                Status:"Closed",
                ContentType:"Service Requests",
                AssignedTeamPerson:"Infrastructure & Secuirty:Kristine Bjorkman",
                Created:moment("2018-09-03T06:37:27Z").format("MMM Do YY"),
                RemainingTime:"7d 10h 38m"
            },
            {
                ID:"17",
                Title:"Equipment Order",
                Priority:"5",
                Prioritycolor:"#456B2B",
                Requester:"Kristine Bjorkman",
                Status:"Converted to Problem",
                ContentType:"Service Requests",
                AssignedTeamPerson:"Kristine Bjorkman",
                Created:moment("2018-09-09T06:37:27Z").format("MMM Do YY"),
                RemainingTime:"9d 10h 38m"
            },
            {
                ID:"110",
                Title:"WiFI is notworking",
                Priority:"2",
                Prioritycolor:"#ED7D31",
                Requester:"Thirumal Kandari",
                Status:"Pending Change",
                ContentType:"Incident",
                AssignedTeamPerson:"Kristine Bjorkman",
                Created:moment("2018-08-03T06:37:27Z").format("MMM Do YY"),
                RemainingTime:"3d 10h 38m"
            },
            {
                ID:"111",
                Title:"WiFI is notworking",
                Priority:"4",
                Prioritycolor:"#7ABC32",
                Requester:"Thirumal Kandari",
                Status:"Pending Change",
                ContentType:"Incident",
                AssignedTeamPerson:"Kristine Bjorkman",
                Created:moment("2019-03-03T06:37:27Z").format("MMM Do YY"),
                RemainingTime:"3d 10h 38m"
            },
            {
                ID:"112",
                Title:"WiFI is notworking",
                Priority:"1",
                Prioritycolor:"#B21E29",
                Requester:"Thirumal Kandari",
                Status:"Pending Change",
                ContentType:"Incident",
                AssignedTeamPerson:"Kristine Bjorkman",
                Created:moment("2018-10-03T06:37:27Z").format("MMM Do YY"),
                RemainingTime:"3d 10h 38m"
            }
        ];
        let error=false;
        return new Promise((resolve,reject)=>{
            setTimeout(()=>{
                if (error) {
                    reject('error'); 
                  } else {
                    resolve(datasource); 
                  }
            },1000);
        });
    }
}
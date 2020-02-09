import { ITeam } from "./ITeam";
import { IUserDetails } from "./IUserDetails";

export interface ITask{
    key?:number;
    Title?:string;
    AssignedTeam?:ITeam;
    AssignedPerson?:IUserDetails;
    StartDate?:string;
    DueDate?:string;
    PercentComplete?:number;
    TaskStatus?:string;
    Description?:string;
}
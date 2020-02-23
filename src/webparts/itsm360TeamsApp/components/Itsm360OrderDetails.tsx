import * as React from 'react';
import { Table, Button } from 'antd';
import { sharepointservice } from '../service/sharepointservice';

export interface IItsm360OrderDetailsProps {
    spservice: sharepointservice;
    orderdetails?: any[];
}

export interface IItsm360OrderDetailsState {
    orders?:any[];
    islicense?:boolean;
}

export class Itsm360OrderDetails extends React.Component<IItsm360OrderDetailsProps, IItsm360OrderDetailsState>{

    constructor(props: IItsm360OrderDetailsProps) {
        super(props);
        this.state = {
            orders:[],
            islicense:false
        };
    }

    public componentDidMount(){
        const { orderdetails,spservice } = this.props;
        let orderdata=[];
        let islicense=false;
        console.log("orderdetails: ",orderdetails);
        orderdetails.forEach((item)=>{
            if(item.isLicense==true){
                islicense=true;
                const order={
                    key:item.ID,
                    ID:item.ID,
                    Title:item.Title,
                    Quantity:item.Quantity,
                    Price:item.Price,
                    Available:item.AwilableAssetsQuantity,
                    Assigned:item.AssignedAssetsQuantity
                };
                orderdata.push(order);
            }else{
                let acount=item.AwilableAssetsQuantity;
                if(typeof item.AwilableAssetsQuantity=="undefined"){
                    spservice.getAssetsAvailability(item.Title).then((lcount)=>{
                        debugger;
                        if(lcount>0){
                            let x=this.state.orders;
                            const removeindex=x.map(y=> y.ID).indexOf(item.ID);
                            const lo={
                                key:item.ID,
                                ID:item.ID,
                                ImageUrl:x[removeindex].ImageUrl,
                                Title:x[removeindex].Title,
                                Quantity:x[removeindex].Quantity,
                                Price:x[removeindex].Price,
                                Available:lcount,
                                Assigned:x[removeindex].Assigned
                            };
                            x.splice(removeindex,1);
                            x.push(lo);
                            this.setState({orders:x});
                        }
                    });
                }
                const order={
                    key:item.ID,
                    ID:item.ID,
                    ImageUrl:item.ImageUrl,
                    Title:item.Title,
                    Quantity:item.Quantity,
                    Price:item.Price,
                    Available:acount,
                    Assigned:typeof item.AssignedAssets!="undefined"?item.AssignedAssets.length:0
                };
                orderdata.push(order);
            }
        });
        this.setState({orders:orderdata,islicense:islicense});
    }

    public render(): React.ReactElement<IItsm360OrderDetailsProps> {
        const {islicense,orders}=this.state;
        const columns = [
            {
                title: '',
                dataIndex: 'ImageUrl',
                key: 'ImageUrl',
                render: url => <img src={url} style={{ width: "100px" }} />,
            },
            {
                title: 'Title',
                dataIndex: 'Title'
            },
            {
                title: 'Quantity',
                dataIndex: 'Quantity'
            },
            {
                title: 'Price',
                dataIndex: 'Price'
            },
            {
                title: 'Available',
                dataIndex: 'Available'
            },
            {
                title: 'Assigned',
                dataIndex: 'Assigned'
            },
        // {
        //     title:'Action',
        //     key:'action',
        //     render:(record)=>(<span>
        //         {record.Available>0 && record.Assigned<=0?
        //             <Button>Assign</Button>:
        //             <Button disabled>Assign</Button>
        //         }
        //         </span>)
        // }
    ];
        const lcolumns=[
            {
                title: 'Title',
                dataIndex: 'Title'
            },
            {
                title: 'Quantity',
                dataIndex: 'Quantity'
            },
            {
                title: 'Price',
                dataIndex: 'Price'
            },
            {
                title: 'Available',
                dataIndex: 'Available'
            },
            {
                title: 'Assigned',
                dataIndex: 'Assigned'
            }];
        return (
            <div>
               {!islicense?<Table size="small" columns={columns} dataSource={orders} pagination={false} />:<Table size="small" columns={lcolumns} dataSource={orders} pagination={false}/>} 
            </div>
        );
    }
}
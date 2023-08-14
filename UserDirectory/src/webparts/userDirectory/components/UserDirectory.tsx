import * as React from 'react';
import {useEffect} from 'react/cjs/react.production.min';
import { useState } from 'react/cjs/react.production.min';
import styles from './UserDirectory.module.scss';
import { IUserDirectoryProps } from './IUserDirectoryProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { spservices } from "../services/spservice";
import { ISPServices } from "../services/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../services/spMockServices";
import { PersonaCard } from "../components/PersonaCard/PersonaCard";
import Paging from '../components/Pagination/Paging';
import {
  Spinner,
  SpinnerSize,
  SearchBox,
  Icon,
  Label,
  Dropdown,
  IDropdownOption
} from "office-ui-fabric-react";
const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');

var startItem;
var endItem;
var alllocQry;
var excludeItems;

export interface MyStates{
  users: any;
  filterUser:any;
  searchString: string;
  searchText: string;
  isLoading: boolean;
  errorMessage: string;
  hasError: boolean;
  pageno:any;
  currentPageNo:any;
  timeout:any;
}

const orderOptions: IDropdownOption[] = [
  { key: "FirstName", text: "First Name" },
  { key: "LastName", text: "Last Name" },
  { key: "Department", text: "Department" },
  { key: "Location", text: "Location" },
  { key: "JobTitle", text: "Job Title" },
  { key: "WorkPhone", text: "Work Phone" },
  { key: "MobilePhone", text: "Mobile Phone" }
];

export default class UserDirectory extends React.Component<IUserDirectoryProps, MyStates> {
  private services: ISPServices = null;
constructor(props){
  super(props);
  this.state={
    users: [],
    filterUser:[],
    searchString: "LastName",
    searchText: "",
    isLoading: true,
    errorMessage: "",
    hasError: false,
    pageno:1,
    currentPageNo:1,
    timeout:0
  };
  debugger;
  if (Environment.type === EnvironmentType.Local) {
    this.services = new spMockServices();
  } else {
  this.services = new spservices(this.props.context);
  }
  this.searchUsers = this.searchUsers.bind(this);
}

public async componentDidMount() {
  debugger;
 await this.searchUsers("A");
}


private async searchUsers(searchText:string){
 searchText = searchText.trim().length > 0 ? searchText : "A";
  var pageSize=this.props.pageSize ? this.props.pageSize : 10;
  this.setState({isLoading:true,currentPageNo:1});
  try{

    if (searchText.length > 0) {
      let searchProps: string[] = this.props.searchProps && this.props.searchProps.length > 0 ?
          this.props.searchProps.split(',') :  ['FirstName', 'LastName', 'WorkEmail', 'Department','WorkPhone','MobilePhone'];
          //this.state.searchString]; //
      let qryText: string = '(';
      let finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;

          searchProps.map((srchprop, index) => {
              if (index == searchProps.length - 1)
                  qryText += `${srchprop}:${finalSearchText}*`;
              else qryText += `${srchprop}:${finalSearchText}* OR `;
          }); 
          qryText +=')';
          let specfiLocation=this.props.specficLoc.trim().toString(); 
          if(specfiLocation!="*"){
            if(specfiLocation.length>0)
            {
              var alllocation = this.props.specficLoc.split(",");
              alllocQry= " AND (";
              // for (let i = 0; i < alllocation.length; i++) {
              //   if(i == alllocation.length-1) {
              //     alllocQry += 'OfficeNumber:'+alllocation[i]+"*"; }
              //    else {
              //     alllocQry += 'OfficeNumber:'+alllocation[i]+"* OR ";
              //   }
              // }
              // qryText +=alllocQry+")";
              qryText += await this.addLocQry(alllocation);
            }
          }
          excludeItems=this.props.exclude;
          console.log(excludeItems);
          const users = await this.services.searchUsersNew('', qryText,excludeItems, false);  
          let searUserDetails=users && users.PrimarySearchResults? users.PrimarySearchResults: null;
          
          startItem = ((1 - 1) * pageSize);
          endItem = 1 * pageSize;
          searUserDetails=this.getSortUsersDetails(searUserDetails,this.state.searchString);
          let filItems = slice(searUserDetails, startItem, endItem);
          this.setState({ users:searUserDetails,filterUser:filItems,
                isLoading: false,
                errorMessage: "",
                hasError: false,pageno:pageSize,currentPageNo:1});
  }
  }
  catch(error){
    console.log("Error Details: "+error);
  }
}

private addLocQry(alllocation) {
 // let result = '';
  //alllocQry= " AND (OfficeNumber:${this.props.specficLoc}*)"
  //let s = alllocation;
 // qryText +=` AND (OfficeNumber:${this.props.specficLoc}*)`;
// var locationQry = this.props.specficLoc
  alllocQry= " AND (";
  for (let i = 0; i < alllocation.length; i++) {
    if(i == alllocation.length-1) {
      alllocQry += 'OfficeNumber:'+alllocation[i]+"*"; }
     else {
      alllocQry += 'OfficeNumber:'+alllocation[i]+"* OR ";
    }
  }
 // console.log(alllocQry+")");
  return alllocQry+")";
}


private searchChanged(e){
//console.log(this.state.searchText);
 // var searchText=e.currentTarget.value;
var searchText=this.state.searchText;
//  console.log(e);
 // setTimeout(async() => {
    this.searchUsers(searchText);
 //   console.log(searchText);
 // }, 1000);
 // this.setState({searchText:searchText});
}

private onPageUpdate(pageno?:number){
  
  var currentPge = (pageno) ? pageno : this.state.currentPageNo;
  startItem = ((currentPge - 1) * this.state.pageno);
  endItem = currentPge * this.state.pageno;
  let filItems = slice(this.state.users, startItem, endItem);
  this.setState({currentPageNo:currentPge,filterUser:filItems});
}

private sortPeople(sortField: string){
  let _users = this.state.users;
  _users = this.getSortUsersDetails(_users,sortField);
  /*_users = _users.sort((a: any, b: any) => {
    switch (sortField) {
      // Sorte by FirstName
      case "FirstName":
        const aFirstName = a.FirstName ? a.FirstName : "";
        const bFirstName = b.FirstName ? b.FirstName : "";
        if (aFirstName.toUpperCase() < bFirstName.toUpperCase()) {
          return -1;
        }
        if (aFirstName.toUpperCase() > bFirstName.toUpperCase()) {
          return 1;
        }
        return 0;
        break;
      // Sort by LastName
      case "LastName":
        const aLastName = a.LastName ? a.LastName : "";
        const bLastName = b.LastName ? b.LastName : "";
        if (aLastName.toUpperCase() < bLastName.toUpperCase()) {
          return -1;
        }
        if (aLastName.toUpperCase() > bLastName.toUpperCase()) {
          return 1;
        }
        return 0;
        break;
      // Sort by Location
      case "Location":
        const aBaseOfficeLocation = a.OfficeNumber
          ? a.OfficeNumber
          : "";
        const bBaseOfficeLocation = b.OfficeNumber
          ? b.OfficeNumber
          : "";
        if (
          aBaseOfficeLocation.toUpperCase() <
          bBaseOfficeLocation.toUpperCase()
        ) {
          return -1;
        }
        if (
          aBaseOfficeLocation.toUpperCase() >
          bBaseOfficeLocation.toUpperCase()
        ) {
          return 1;
        }
        return 0;
        break;
      // Sort by JobTitle
      case "JobTitle":
        const aJobTitle = a.JobTitle ? a.JobTitle : "";
        const bJobTitle = b.JobTitle ? b.JobTitle : "";
        if (aJobTitle.toUpperCase() < bJobTitle.toUpperCase()) {
          return -1;
        }
        if (aJobTitle.toUpperCase() > bJobTitle.toUpperCase()) {
          return 1;
        }
        return 0;
        break;
      // Sort by Department
      case "Department":
        const aDepartment = a.Department ? a.Department : "";
        const bDepartment = b.Department ? b.Department : "";
        if (aDepartment.toUpperCase() < bDepartment.toUpperCase()) {
          return -1;
        }
        if (aDepartment.toUpperCase() > bDepartment.toUpperCase()) {
          return 1;
        }
        return 0;
        break;
      default:
        break;
    }
  });*/
  let filItems = slice(_users, startItem, endItem);
  this.setState({ users: _users,filterUser:filItems, searchString: sortField });
}
  private getSortUsersDetails(_users: any, sortField: string): any {
    let sortUser;
    sortUser = _users.sort((a: any, b: any) => {
      switch (sortField) {
        // Sorte by FirstName
        case "FirstName":
          const aFirstName = a.FirstName ? a.FirstName : "";
          const bFirstName = b.FirstName ? b.FirstName : "";
          if (aFirstName.toUpperCase() < bFirstName.toUpperCase()) {
            return -1;
          }
          if (aFirstName.toUpperCase() > bFirstName.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by LastName
        case "LastName":
          const aLastName = a.LastName ? a.LastName : "";
          const bLastName = b.LastName ? b.LastName : "";
          if (aLastName.toUpperCase() < bLastName.toUpperCase()) {
            return -1;
          }
          if (aLastName.toUpperCase() > bLastName.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by Location
        case "Location":
          const aBaseOfficeLocation = a.OfficeNumber
            ? a.OfficeNumber
            : "";
          const bBaseOfficeLocation = b.OfficeNumber
            ? b.OfficeNumber
            : "";
          if (
            aBaseOfficeLocation.toUpperCase() <
            bBaseOfficeLocation.toUpperCase()
          ) {
            return -1;
          }
          if (
            aBaseOfficeLocation.toUpperCase() >
            bBaseOfficeLocation.toUpperCase()
          ) {
            return 1;
          }
          return 0;
          break;
        // Sort by JobTitle
        case "JobTitle":
          const aJobTitle = a.JobTitle ? a.JobTitle : "";
          const bJobTitle = b.JobTitle ? b.JobTitle : "";
          if (aJobTitle.toUpperCase() < bJobTitle.toUpperCase()) {
            return -1;
          }
          if (aJobTitle.toUpperCase() > bJobTitle.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        // Sort by Department
        case "Department":
          const aDepartment = a.Department ? a.Department : "";
          const bDepartment = b.Department ? b.Department : "";
          if (aDepartment.toUpperCase() < bDepartment.toUpperCase()) {
            return -1;
          }
          if (aDepartment.toUpperCase() > bDepartment.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
           // Sort by Work Phone
        case "WorkPhone":
          const aWorkPhone = a.WorkPhone ? a.WorkPhone : "";
          const bWorkPhone = b.WorkPhone ? b.WorkPhone : "";
          if (aWorkPhone.toUpperCase() < bWorkPhone.toUpperCase()) {
            return -1;
          }
          if (aWorkPhone.toUpperCase() > bWorkPhone.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
          // Sort by Mobile No
        case "MobilePhone":
          const aMobilePhone = a.MobilePhone ? a.MobilePhone : "";
          const bMobilePhone = b.MobilePhone ? b.MobilePhone: "";
          if (aMobilePhone.toUpperCase() < bMobilePhone.toUpperCase()) {
            return -1;
          }
          if (aMobilePhone.toUpperCase() > bMobilePhone.toUpperCase()) {
            return 1;
          }
          return 0;
          break;
        default:
          break;
      }
    });
    return sortUser;
  }


  public render(): React.ReactElement<IUserDirectoryProps> {
    debugger;
    const color = this.props.context.microsoftTeams ? "white" : "";
    const diretoryGrid =
      this.state.filterUser && this.state.filterUser.length > 0
        ? this.state.filterUser.map((user: any) => {
          return (
            <PersonaCard
              context={this.props.context}
              profileProperties={{
                DisplayName: user.PreferredName,
                Title: user.JobTitle,
                PictureUrl: user.PictureURL,
                Email: user.WorkEmail,
                Department: user.Department,
                WorkPhone: user.WorkPhone,
                MobilePhone: user.MobilePhone,
                Location: user.OfficeNumber
                  ? user.OfficeNumber
                  : user.BaseOfficeLocation
              }}
            />
          );
        })
        : [];

  
    return (
        <div>
        <SearchBox
            placeholder="Search Users"
            styles={{
              root: {
                minWidth: 180,
                maxWidth: 300,
                marginLeft: "auto",
                marginRight: "auto",
                marginBottom: 25
              }
            }}
            value={this.state.searchText}
            onChange={(event, value) => this.setState({ searchText: value })}
            onSearch={this.searchChanged.bind(this)}
            onClear={ ()=>{console.log("hisde");this.setState({ searchText:"A*" });this.searchUsers("A");}}
           //   onChange={this.searchChanged.bind(this)}
          />
           {(!this.state.users || this.state.users.length == 0 )? 
          <div className={styles.noUsers}>
            <Icon
              iconName={"ProfileSearch"}
              style={{ fontSize: "54px", color: color }}
            />
            <Label>
              <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>No users found in directory</span>
            </Label>
          </div>
           :(this.state.isLoading) ? (
            <Spinner size={SpinnerSize.large} label={"searching ..."} />
          ) :
           <div>
         <div className={styles.dropDownSortBy}>
          <Dropdown
                    placeholder="Sort by"
                    label="Sort People by"
                    options={orderOptions}
                    selectedKey={this.state.searchString}
                    onChange={(ev: any, value: IDropdownOption) => {
                      console.log(this.state.searchString);
                      console.log(value+"->"+value.key.toString());
                      this.sortPeople(value.key.toString());
                    }}
                    styles={{ dropdown: { width: 200 } }}
                  />
             </div>
           <div>{diretoryGrid}</div>
           <div style={{ width: '100%', display: 'inline-block' }}>
              <Paging
                  totalItems={this.state.users.length}
                  itemsCountPerPage={this.state.pageno}
                  onPageUpdate={this.onPageUpdate.bind(this)}
                  currentPage={this.state.currentPageNo} />
            </div>
            </div>
           }
        </div>
    );
  }
}

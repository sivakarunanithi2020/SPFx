
export interface ISPServices {

    searchUsers(searchString: string, searchFirstName: boolean);
    searchUsersNew(searchString: string, srchQry: string, excludeItems:string[], isInitialSearch: boolean, pageNumber?: number);

}
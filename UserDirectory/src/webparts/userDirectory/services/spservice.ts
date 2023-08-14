import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/search";
import { sp } from '@pnp/sp';
import { SearchQueryBuilder, SearchResults, ISearchQuery,SortDirection } from "@pnp/sp/search";


export class spservices  {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async searchUsers(searchString: string, searchFirstName: boolean): Promise<SearchResults> {
        debugger;
        const _search = !searchFirstName ? `LastName:${searchString}*` : `FirstName:${searchString}*`;
        const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId"];
        try {
            if (!searchString) return undefined;
            let users = await sp.search(<ISearchQuery>{
                Querytext: _search,
                RowLimit: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
            });
            console.log(_search);
            console.log(users);
            return users;
        } catch (error) {
            Promise.reject(error);
        }
    }

    public async searchUsersNew(searchString: string, srchQry: string, excludeItems:string[], isInitialSearch: boolean, pageNumber?: number): Promise<SearchResults> {
        let qrytext: string = '';
        if (isInitialSearch) qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
        else {
            if (srchQry) qrytext = srchQry;
            else {
                if (searchString) qrytext = searchString;
            }
            if (qrytext.length <= 0) qrytext = `*`;
            console.log(qrytext.length);
        }
        console.log(qrytext.length);
        console.log(qrytext);
        const searchProperties: string[] = [ "FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId"];
        //const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills","BaseOfficeLocation"];
        let RefinedProperties:string[]=[""];
        if(excludeItems!=null)
        {
            RefinedProperties = excludeItems.toString().split(','); // ["UK","_srv"];
        }
        let RefinedQry=[];
        for (let i = 0; i < RefinedProperties.length; i++) {
           // if(i == RefinedProperties.length-1) {
                RefinedQry[i] = "PreferredName:not(('"+RefinedProperties[i]+"*'))"; }
             //else {
             //   RefinedQry[i] = "PreferredName:not(('"+RefinedProperties[i]+"'*)) OR ";
            //}
          //}
          console.log(RefinedQry);
          console.log(qrytext);
        try {
            ///sp.search
            let users = await sp.search(<ISearchQuery>{
                Querytext: qrytext,
                RowLimit: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                RefinementFilters: RefinedQry,
                //["PreferredName:not(('UK*'))","PreferredName:not(('_srv*'))"],
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
            });
            if (users && users.PrimarySearchResults.length > 0) {
                for (let index = 0; index < users.PrimarySearchResults.length; index++) {
                    let user: any = users.PrimarySearchResults[index];
                    if (user.PictureURL) {
                        user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        users.PrimarySearchResults[index] = user;
                    }
                }
            }  
            console.log(users);
           // console.log(qrytext);
            return users;

        } catch (error) {
            Promise.reject(error);
        }
    }
}
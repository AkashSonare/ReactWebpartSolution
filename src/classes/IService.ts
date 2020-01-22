import {
    SPHttpClient,
    SPHttpClientResponse,MSGraphClient
} from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PageContext } from '@microsoft/sp-page-context';
import {IStateCard} from '../classes/IState';

export class IServiceClass {        
    private _spHttpClient: SPHttpClient;    
    private _pageContext: PageContext;    
    private _currentWebUrl: string;
    public getWelcomeMessageDetails(context: WebPartContext, url: string): Promise<IStateCard> {
        var url = `${context.pageContext.web.absoluteUrl}${url}`;        
        return context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
            return response.json();
        });
    }

    public getCurrentUserId(context: any, userid: string): Promise<number> {
        return context.spHttpClient.get(context.pageContext.web.absoluteUrl + "/_api/web/siteusers(@v)?@v='" + encodeURIComponent(userid) + "'", SPHttpClient.configurations.v1)
        .then((responeMaster: SPHttpClientResponse) => {
            return responeMaster.json().then((obj) => {
                return obj.Id;  
            });
        });
    }
}
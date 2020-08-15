import * as React from 'react';
import styles from './MyMails.module.scss';
import { PublicClientApplication, 
          InteractionRequiredAuthError,          
          AuthorizationUrlRequest, 
          AccountInfo} from "@azure/msal-browser";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { IMail } from '../../../model/IMail';
import { IMyMailsProps } from './IMyMailsProps';
import { IMyMailsState } from './IMyMailsState';

export default class MyMails extends React.Component<IMyMailsProps, IMyMailsState> {  
  private myMSALObj: PublicClientApplication;
  private ssoRequest: AuthorizationUrlRequest = {
    scopes: ["https://graph.microsoft.com/Mail.Read"]    
  };

  constructor(props) {
    super(props);

    const msalConfig = {
      auth: {
        authority: `https://login.microsoftonline.com/${this.props.tenantUrl}`,
        clientId: this.props.applicationID,
        redirectUri: this.props.redirectUri
      }
    };
    
    this.myMSALObj = new PublicClientApplication(msalConfig);

    this.state = {
      mails: []
    };
    this.myMSALObj.handleRedirectPromise().then((tokenResponse) => {      
      if (tokenResponse !== null) {
        const access_token = tokenResponse.accessToken;
        this.getMailsFromGraph(access_token).then(mails => {
          this.setState(() => {
            return { mails: mails };
          });      
        });
      } else 
      {
        // In case we would like to directly load data in case of NO redirect:
        // const currentAccounts = this.myMSALObj.getAllAccounts();
        // this.handleLoggedInUser(currentAccounts);
      }      
    }).catch((error) => {
        console.log(error);
        return null;
    });
  }

  public render(): React.ReactElement<IMyMailsProps> {
    let mailElements = this.state.mails.map(mail => {
      return <li>{mail.from} - {mail.subject}</li>;
    });
    return (
      <div className={ styles.myMails }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <a onClick={this.loadMails} className={ styles.button }>
                <span className={ styles.label }>Get mails</span>
              </a>
            </div>
          </div>
          <div className={styles.row}>
            <ul>
              {mailElements}
            </ul>
          </div>
        </div>
      </div>
    );
  }

  private loadMails = () => {
    const accounts = null; //this.myMSALObj.getAllAccounts();
    if (accounts !== null) {
      this.handleLoggedInUser(accounts);
    }
    else {
      this.loginForAccessTokenByMSAL()
      .then((token) => {
        this.getMailsFromGraph(token).then(mails => {
          this.setState(() => {
            return { mails: mails };
          });      
        });
      });
    }    
  }

  private async getMailsFromGraph(accessToken: string): Promise<any> {
    if (accessToken !== null) {
      const graphMailEndpoint: string = "https://graph.microsoft.com/v1.0/me/messages";
      return this.props.httpClient
        .get(graphMailEndpoint, HttpClient.configurations.v1,
          {
            headers: [
              ['Authorization', `Bearer ${accessToken}`]
            ]
          })
        .then((res: HttpClientResponse): Promise<any> => {
          return res.json();
        })
        .then((response: any) => {
          console.log(response);
          let mails: IMail[] = [];
          response.value.forEach((m) => {
            mails.push({from: m.from.emailAddress.address, subject: m.subject});
          });
          return mails;
        });
      }
      else {
        console.log("Error retrieving token");
        return [];
      }
  }

  private async loginForAccessTokenByMSAL(): Promise<string> {
    this.ssoRequest.loginHint = this.props.userMail;  
    return this.myMSALObj.ssoSilent(this.ssoRequest).then((response) => {
      return response.accessToken;  
    }).catch((silentError) => {  
        console.log(silentError);
        if (silentError instanceof InteractionRequiredAuthError) {
          return this.myMSALObj.loginPopup(this.ssoRequest)
          .then((response) => {
            return response.accessToken;
          }) 
          .catch(popupError => {
            if (popupError.message.indexOf('popup_window_error') > -1) { // Popups are blocked
              return this.redirectLogin(this.ssoRequest);
            }            
          });
        } else {
            return null;
        }
    });  
  }

  private handleLoggedInUser(currentAccounts: AccountInfo[]) {
    let accountObj = null;
    if (currentAccounts === null) {
      // No user signed in
      return;
    } else if (currentAccounts.length > 1) {
        // More than one user is authenticated, get current one 
        accountObj = this.myMSALObj.getAccountByUsername(this.props.userMail);
    } else {
        accountObj = currentAccounts[0];
    }
    if (accountObj === null) {
      this.acquireAccessToken(this.ssoRequest, accountObj)
      .then((accessToken) => {
        this.getMailsFromGraph(accessToken).then(mails => {
          this.setState(() => {
            return { mails: mails };
          });      
        });
      });
    }    
  }

  private async acquireAccessToken(ssoRequest: AuthorizationUrlRequest, account: AccountInfo): Promise<string> {
    const accessTokenRequest = {
      scopes: ssoRequest.scopes,
      account: account
    };
    return this.myMSALObj.acquireTokenSilent(accessTokenRequest).then((val) => {            
      return val.accessToken;
    }).catch((errorinternal) => {  
      console.log(errorinternal);  
      return null;
    }); 
  }

  private redirectLogin(ssoRequest: AuthorizationUrlRequest): Promise<string> {
    try {      
      this.myMSALObj.loginRedirect(ssoRequest)
        .then(() => {
          return Promise.resolve('');
        });
      
    } catch (err) {
        console.log(err);
        return null;
    }
  }
}

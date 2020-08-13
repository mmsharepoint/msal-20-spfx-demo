import * as React from 'react';
import styles from './MyMails.module.scss';
import { PublicClientApplication, 
          InteractionRequiredAuthError,          
          AuthenticationResult,
          AuthorizationUrlRequest } from "@azure/msal-browser";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { IMail } from '../../../model/IMail';
import { IMyMailsProps } from './IMyMailsProps';
import { IMyMailsState } from './IMyMailsState';

export default class MyMails extends React.Component<IMyMailsProps, IMyMailsState> {  
  private myMSALObj: PublicClientApplication;

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
      let accountObj = null;
      if (tokenResponse !== null) {
        const access_token = tokenResponse.accessToken;
        this.getMailsByMSAL(access_token).then(mails => {
          this.setState(() => {
            return { mails: mails };
          });      
        });
      } else 
      {
        const currentAccounts = this.myMSALObj.getAllAccounts();
        if (currentAccounts === null) {
            // No user signed in
            return;
        } else if (currentAccounts.length > 1) {
            // More than one user signed in, find desired user with 
            accountObj = this.myMSALObj.getAccountByUsername(this.props.userMail);
        } else {
            accountObj = currentAccounts[0];
        }
        // acquireAccessToken with request and request.account = accountObject
        // then go on ...
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
    this.getAccessTokenByMSAL()
      .then((token) => {
        this.getMailsByMSAL(token).then(mails => {
          this.setState(() => {
            return { mails: mails };
          });      
        });
      });
  }

  private async getMailsByMSAL(accessToken: string): Promise<any> {
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

  private async getAccessTokenByMSAL(): Promise<string> {  
    const ssoRequest = {
      scopes: ["https://graph.microsoft.com/Mail.Read"],
      loginHint: this.props.userMail
    };
    const accounts = this.myMSALObj.getAllAccounts();
    return this.myMSALObj.ssoSilent(ssoRequest).then((response) => {
      return this.acquireAccessToken(ssoRequest, response);  
    }).catch((error) => {  
        console.log(error);
        if (error instanceof InteractionRequiredAuthError) {
          return this.myMSALObj.loginPopup(ssoRequest)
          .then((response) => {
            return this.acquireAccessToken(ssoRequest, response);
          }) 
          .catch(error => {
            if (error.message.indexOf('popup_window_error') > -1) { // Popups are blocked
              return this.redirectLogin(ssoRequest);
            }            
          });
        } else {
            return null;
        }
    });  
  }

  private async acquireAccessToken(ssoRequest: AuthorizationUrlRequest, authResult: AuthenticationResult): Promise<string> {
    const accessTokenRequest = {
      scopes: ssoRequest.scopes,
      account: authResult.account
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

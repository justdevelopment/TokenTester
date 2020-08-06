import * as React from 'react';
import styles from './TokenTest.module.scss';
import { MSGraphClient } from '@microsoft/sp-http';
import jwt from 'jsonwebtoken';

export interface ITokenTestProps {
  context: any;
}

export interface ITokenTestState {
  result: any;
  timestamp: string;
  isSuccess: boolean;
  token: any;
  tokenExpiration: any;
  expiration: any;
}


export default class TokenTest extends React.Component<ITokenTestProps, ITokenTestState> {
  private client: MSGraphClient;

  constructor(props) {
    super(props);

    this.state = {
      isSuccess: true,
      result: null,
      timestamp: null,
      token: null,
      tokenExpiration: null,
      expiration: null
    };
  }

  public getStorageByPartialKey(partialKeys) {
    let result = [];
    for (var i = sessionStorage.length - 1; i >= 0; --i) {
      var key = sessionStorage.key(i);
      partialKeys.forEach(partialKey => {
        if (key.indexOf(partialKey) >= 0) {
          result.push(sessionStorage.getItem(key));
        }
      });
    }

    return result;
  }

  private doCall() {
    this.client.api("/me")
      .select("id,mail")
      .get((error, response) => {
        try {
          let graphKey = ["adal.access.token.key|https://graph.microsoft.com", "adal.access.token.keyhttps://graph.microsoft.com"];
          let expirationKey = ["adal.expiration.key|https://graph.microsoft.com", "adal.expiration.keyhttps://graph.microsoft.com"];

          let tokens = this.getStorageByPartialKey(graphKey);
          let expiration = this.getStorageByPartialKey(expirationKey);

          let expirationTime = Number(expiration[0]);
          let token = jwt.decode(tokens[0]);
          let tokenExp = token["exp"];

          this.setState({
            timestamp: new Date().toISOString(),
            token: JSON.stringify(token, null, 4),
            tokenExpiration: tokenExp,
            expiration: expirationTime
          });

          if (error)
            return this.setState({
              isSuccess: false,
              result: JSON.stringify(error, null, 4)
            });

          this.setState({
            isSuccess: true,
            result: JSON.stringify(response, null, 4)
          });
        } catch (err) {
          return this.setState({
            isSuccess: false,
            result: JSON.stringify(err, null, 4)
          });
        }
      });
  }

  public componentDidMount() {
    this.props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient): void => {
        this.client = client;
        this.doCall();
        setInterval(this.doCall.bind(this), 300000); // refresh every 15 min
      });
  }

  public render(): React.ReactElement<ITokenTestProps> {
    return (
      <div className={styles.tokenTest}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Testing graph token refresh</p>
              <p className={styles.description}>
                Result of last call: {this.state.isSuccess ? "Success!" : "Failure: " + this.state.result} <br />Call done on: {this.state.timestamp}</p>
              <p className={styles.description}>
                Expiration: {this.state.expiration} = {new Date(this.state.expiration).toISOString()}<br />
                Token Expiration: {this.state.tokenExpiration} = {new Date(this.state.tokenExpiration * 1000).toISOString()}<br />
              </p>
              <p className={styles.description}>
                Full Token: <pre>{this.state.token}</pre></p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

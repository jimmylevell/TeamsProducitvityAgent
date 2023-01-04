// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React from 'react';
import './App.css';
import { TeamsFx } from "@microsoft/teamsfx";
import { Button } from "@fluentui/react-northstar"
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { CacheService } from '@microsoft/mgt';

class DevOps extends React.Component {
  constructor(props) {
    super(props);
    CacheService.clearCaches();
    this.state = {
      showLoginPage: undefined,
      token: undefined,
      workItems: []
    }
  }
  async componentDidMount() {

    /*Define scope for the required permissions*/
    this.scope = [
      "499b84ac-1321-427f-aa17-267ca6975798/.default"
    ];

    /*Initialize TeamsFX provider*/
    this.teamsfx = new TeamsFx();
    const provider = new TeamsFxProvider(this.teamsfx, this.scope)
    Providers.globalProvider = provider;

    /*Check if consent is needed*/
    let consentNeeded = false;
    try {
      let token = await this.teamsfx.getCredential().getToken(this.scope)
      this.setState({
        token: token
      }, () => {
        this.loadGraphData(this.state.token)
      });
    } catch (error) {
      consentNeeded = true;
    }
    this.setState({
      showLoginPage: consentNeeded
    });
    Providers.globalProvider.setState(consentNeeded ? ProviderState.SignedOut : ProviderState.SignedIn);
    return consentNeeded;
  }

  async loadGraphData(token) {
    const devops = "https://dev.azure.com/levellch/kanban-cleanup/_apis/work/teamsettings/iterations?api-version=7.0&$timeframe=current"

    const response = await fetch(devops, {
      method: 'GET',
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Allow-control-allow-credentials': 'true',
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token.token}`,
      }
    });

    const data = await response.json();

    let currentSprint = data.value.find(sprint => sprint.attributes.timeFrame === "current")

    const workItemsGet = await fetch(currentSprint.url + "/workitems?api-version=7.0", {
      method: 'GET',
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Allow-control-allow-credentials': 'true',
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token.token}`,
      }
    });

    let response2 = await workItemsGet.json();

    let workItems = response2.workItemRelations.map(item => {
      return { id: item.target.id, url: item.target.url }
    })

    let fullWorkItems = []
    await workItems.forEach(async (item) => {
      const workItemsGet = await fetch(item.url + "?api-version=7.0", {
        method: 'GET',
        headers: {
          'Access-Control-Allow-Origin': '*',
          'Allow-control-allow-credentials': 'true',
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token.token}`,
        }
      });
      fullWorkItems.push(await workItemsGet.json())
    })

    this.setState({
      workItems: fullWorkItems
    })

  }

  async loginBtnClick() {
    try {
      await this.teamsfx.login(this.scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
      this.setState({
        showLoginPage: false
      });
    } catch (err) {
      if (err.message?.includes("CancelledByUser")) {
        const helpLink = "https://aka.ms/teamsfx-auth-code-flow";
        err.message +=
          "\nIf you see \"AADSTS50011: The reply URL specified in the request does not match the reply URLs configured for the application\" " +
          "in the popup window, you may be using unmatched version for TeamsFx SDK (version >= 0.5.0) and Teams Toolkit (version < 3.3.0) or " +
          `cli(version < 0.11.0).Please refer to the help link for how to fix the issue: ${helpLink}`;
      }
      alert("Login failed: " + err);
      return;
    }
  }
  render() {
    console.log(this.state.workItems)
    return (
      <div>
        {
          this.state.showLoginPage === false &&

          <div>
            <div className="features">
              <div className="header">
                <div className="title">
                  <div className="row">
                    <div className="column">
                      <h3>Planned Items</h3>
                    </div>
                    <div className="column">
                      <h3>Doing Items</h3>
                    </div>
                    <div className="column">
                      <h3>Done</h3>
                    </div>
                  </div>
                </div>
              </div>
              <div className="row content">
                <div className="column mgt-col">
                  {this.state.workItems.map(item => {

                    return <div key={item.id} className="card">
                      <div className="card-body">
                        <h5 className="card-title">cbcfb{item.fields["System.Title"]}</h5>
                        <p className="card-text">{item.fields["System.Description"]}</p>
                      </div>
                    </div>

                  })
                  }
                </div>
              </div>
            </div>
          </div>
        }
        {
          this.state.showLoginPage === true &&
          <div className="auth">
            <h3>Welcome to the Producitivty Coach!</h3>
            <p>Please click on "Start" and consent permissions to use the app.</p>
            <Button primary onClick={() => this.loginBtnClick()}>Start</Button>
          </div>
        }
      </div>
    );
  }
}
export default DevOps;

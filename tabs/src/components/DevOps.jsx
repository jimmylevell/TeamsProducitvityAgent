// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useState, useEffect } from 'react';
import './App.css';
import { TeamsFx } from "@microsoft/teamsfx";
import { Button, Flex, Segment, Image, Grid } from "@fluentui/react-northstar"
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { TeamsFxProvider } from '@microsoft/mgt-teamsfx-provider';
import { CacheService } from '@microsoft/mgt';

import Issue from './Issue';

export default function DevOps(props) {
  const [showLoginPage, setShowLoginPage] = useState(undefined);
  const [token, setToken] = useState(undefined);
  const [workItems, setWorkItems] = useState([]);

  const teamsfx = new TeamsFx();
  const scope = [
    "499b84ac-1321-427f-aa17-267ca6975798/.default"
  ];
  const provider = new TeamsFxProvider(teamsfx, scope)
  Providers.globalProvider = provider;

  CacheService.clearCaches();

  useEffect(() => {
    let consentNeeded = false;
    try {
      teamsfx.getCredential().getToken(scope)
        .then(token => {
          if (token && token.token) {

            setToken(token)
            loadGraphData(token)
          }
        })
    } catch (error) {
      consentNeeded = true;
    }
    setShowLoginPage(consentNeeded)
    Providers.globalProvider.setState(consentNeeded ? ProviderState.SignedOut : ProviderState.SignedIn);
  }, [])

  const load = async (url, method, token) => {
    const response = await fetch(url, {
      method: method,
      headers: {
        'Access-Control-Allow-Origin': '*',
        'Allow-control-allow-credentials': 'true',
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token.token}`,
      }
    });

    return await response.json();
  }

  const loadGraphData = async (token) => {
    const devops = "https://dev.azure.com/levellch/kanban-cleanup/_apis/work/teamsettings/iterations?api-version=7.0&$timeframe=current"
    const sprints = await load(devops, "GET", token);

    let currentSprint = sprints.value.find(sprint => sprint.attributes.timeFrame === "current")

    const workItemsData = await load(currentSprint.url + "/workitems?api-version=7.0", "GET", token)

    let workItems = workItemsData.workItemRelations.map(item => {
      return { id: item.target.id, url: item.target.url }
    })

    let fullWorkItems = []
    await Promise.all(workItems.map(async (item) => {
      const workItemsGet = await load(item.url + "?api-version=7.0&$expand=All", "GET", token);
      fullWorkItems.push(workItemsGet)
    }))

    setWorkItems(fullWorkItems)
  }

  const loginBtnClick = async () => {
    try {
      await teamsfx.login(this.scope);
      Providers.globalProvider.setState(ProviderState.SignedIn);
      setShowLoginPage(false);
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

  console.log(workItems)
  return (
    <>
      {
        showLoginPage === false &&
        <>
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
                  {workItems.map(item => {
                    if (item.fields["System.BoardColumn"] === "To Do") {
                      return <Issue key={item.id} item={item} load={load} token={token} />
                    } else {
                      return null
                    }
                  })
                  }
                </div>
                <div className="column mgt-col">
                  {workItems.map(item => {
                    if (item.fields["System.BoardColumn"] === "Doing") {
                      return <Issue key={item.id} item={item} load={load} token={token} />
                    } else {
                      return null
                    }
                  })
                  }
                </div>
                <div className="column mgt-col">
                  {workItems.forEach(item => {
                    if (item.fields["System.BoardColumn"] === "Done") {
                      return <Issue key={item.id} item={item} load={load} token={token} />
                    } else {
                      return null
                    }
                  })
                  }
                </div>
              </div>
            </div>
          </div>
        </>
      }
      {
        showLoginPage === true &&
        <div className="auth">
          <h3>Welcome to the Producitivty Coach!</h3>
          <p>Please click on "Start" and consent permissions to use the app.</p>
          <Button primary onClick={() => loginBtnClick()}>Start</Button>
        </div>
      }
    </>
  );
}

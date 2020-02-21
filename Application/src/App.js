import config from "./Config";
import MessageCardList from "./MessageCardList";
import {
  getChannel,
  getChannelsOfTeam,
  getJoinedTeams,
  getMessagesOfChannel,
  getRepliesOfMessage,
  getTeam,
  getUsers,
  postMessage
} from "./GraphService";
import { setupUserAgentApplication, logout } from "./GraphFunction";
import { getQueryParams } from "./UrlUtil";

// yarn add @material-ui/core
// yarn add @material-ui/icons
import Button from "@material-ui/core/Button";

// yarn add react-json-tree
import JSONTree from "react-json-tree";

// yarn add react-dropdown-tree-select
import DropdownTreeSelect from "react-dropdown-tree-select";
import "react-dropdown-tree-select/dist/styles.css";

// yarn add material-table
import React, { Component } from "react";
import MaterialTable from "material-table";

// yarn add react-toastify
// import React, { Component } from 'react';
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";

class App extends Component {
  constructor(props) {
    super(props);

    if (config.isDebug) {
      console.log("isDebug");
    }

    this.state = {
      channels: [],
      chatMessageText: "",
      columnsMessageTable: [
        {
          title: "ID",
          field: "id",
          editable: "never"
        },
        {
          title: "DisplayName",
          field: "from.user.displayName",
          editable: "never"
        },
        {
          title: "Content",
          field: "body.content"
        },
        {
          title: "createdDateTime",
          field: "createdDateTime",
          editable: "never"
        }
      ],
      columnsUserTable: [
        {
          title: "ID",
          field: "id",
          editable: "never"
        },
        {
          title: "DisplayName",
          field: "displayName"
        },
        {
          title: "userPrincipalName",
          field: "userPrincipalName"
        },
        {
          title: "mobilePhone",
          field: "mobilePhone"
        }
      ],
      isAuthenticated: false,
      isDropdownTreeSelectDisabled: false,
      messages: [],
      selected: {
        channel: {
          description: "",
          id: "",
          name: ""
        },
        message: { id: "" },
        team: {
          description: "",
          id: "",
          teamName: ""
        }
      },
      user: {},
      users: []
    };

    if (config.isDebug) {
      console.log("this.state");
      console.log(this.state);
    }

    this.ReadGraphData = this.ReadGraphData.bind(this);
    this.ReadGraphMessagesData = this.ReadGraphMessagesData.bind(this);
    this.ReadGraphTeamsData = this.ReadGraphTeamsData.bind(this);
    this.Notify = this.Notify.bind(this);
    this.onTreeChange = this.onTreeChange.bind(this);
    this.onTreeAction = this.onTreeAction.bind(this);
    this.onTreeNodeToggle = this.onTreeNodeToggle.bind(this);
    this.postChatMessage = this.postChatMessage.bind(this);

    try {
      let userAgentApplication = setupUserAgentApplication(this);
      this.ReadGraphData(userAgentApplication);
    } catch (err) {
      if (config.isDebug) {
        console.log("ReadGraphTeamsData() err");
        console.log(err);
      }

      this.Notify(
        "error",
        `エラーが発生しました: ${err.message} : ${err.fileName} : ${err.lineNumber}`
      );
    }
  }

  async ReadGraphData() {
    const params = await getQueryParams(window.location.search);

    // forceパラメータがある場合は強制的にサーバから情報取得
    if (params["force"]) {
      console.log("this.ReadGraphTeamsData(true)");
      await this.ReadGraphTeamsData(true);
      return;
    }

    if (params["teamId"] && params["channelId"]) {
      if (config.isDebug) {
        console.log('params["teamId"]');
        console.log(params["teamId"]);
        console.log('params["channelId"]');
        console.log(params["channelId"]);

        console.log(
          "this.ReadGraphMessagesData();this.ReadGraphTeamsData(false)"
        );
      }
      this.ReadGraphMessagesData(params["teamId"], "", params["channelId"], "");
      await this.ReadGraphTeamsData(false);
      return;
    }

    const loaded = localStorage.getItem("state");
    if (loaded) {
      if (config.isDebug) {
        console.log("loaded");
        console.log(loaded);
        console.log(JSON.parse(loaded));
      }

      this.setState(JSON.parse(loaded));
    }

    if (config.isDebug) {
      console.log("this.ReadGraphTeamsData(false)");
    }
    await this.ReadGraphTeamsData(false);
  }

  async onTreeChange(currentNode, selectedNodes) {
    if (config.isDebug) {
      console.log("onTreeChange:", currentNode, selectedNodes);
    }
    if (selectedNodes.length === 1 && selectedNodes[0]["_depth"] === 2) {
      var splitedvalue = selectedNodes[0].value.split("/");
      var splitedLabel = selectedNodes[0].label.split("/");
      if (config.isDebug) {
        console.log(
          "onTreeChange:",
          this.state.selected.team.id,
          splitedvalue[0],
          this.state.selected.channel.id,
          splitedvalue[1]
        );
      }
      if (
        this.state.selected.team.id !== splitedvalue[0] ||
        this.state.selected.channel.id !== splitedvalue[1]
      ) {
        if (config.isDebug) {
          console.log(
            "表示するチャネル:",
            splitedvalue[0],
            splitedLabel[0],
            splitedvalue[1],
            splitedLabel[1]
          );
        }
        await this.ReadGraphMessagesData(
          splitedvalue[0],
          splitedLabel[0],
          splitedvalue[1],
          splitedLabel[1]
        );
      }
      return;
    }
  }

  onTreeAction = (node, action) => {
    console.log("onTreeAction:", action, node);
  };

  onTreeNodeToggle = currentNode => {
    console.log("onTreeNodeToggle:", currentNode);
  };

  async postChatMessage() {
    if (this.state.chatMessageText) {
      console.log(this.state.chatMessageText);

      // Get the user's accessr token
      var accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });

      postMessage(
        accessToken,
        this.state.selected.team.id,
        this.state.selected.channel.id,
        "undefined" != typeof this.state.selected.message &&
          "undefined" != typeof this.state.selected.message.id
          ? this.state.selected.message.id
          : "",
        this.state.chatMessageText ? this.state.chatMessageText : ""
      );
    }
  }

  Notify = (type, message) => {
    if (config.isDebug) {
      console.log("Notify(" + String(type) + ", " + String(message) + ")");
    }

    var date = new Date();
    if (config.isDebug) {
      console.log("date");
      console.log(date);
    }
    var toastId = date.getTime(); // UNIXTIME(msec)
    if (config.isDebug) {
      console.log("toastId");
      console.log(toastId);
    }

    switch (type) {
      case "info":
        toast.info(message, {
          toastId: toastId
        });
        break;
      case "success":
        toast.success(message, {
          toastId: toastId
        });
        break;
      case "warning":
        toast.warn(message, {
          toastId: toastId
        });
        break;
      case "error":
        toast.error(message, {
          toastId: toastId
        });
        break;
      default:
        toast(message, {
          toastId: toastId
        });
        break;
    }
  };

  async ReadGraphMessagesData(teamId, teamName, channelId, channelName) {
    if (config.isDebug) {
      console.log("ReadGraphMessagesData()");
    }

    this.setState({
      isDropdownTreeSelectDisabled: true
    });

    // Get the user's accessr token
    var accessToken = await window.msal.acquireTokenSilent({
      scopes: config.scopes
    });

    if (teamName === "" || channelName === "") {
      var gotTeam = await getTeam(accessToken, teamId);
      console.log("gotTeam");
      console.log(gotTeam);
      teamName = gotTeam.displayName;

      var gotChannel = await getChannel(accessToken, teamId, channelId);
      console.log("gotChannel");
      console.log(gotChannel);
      channelName = gotChannel.displayName;
    }

    var gotmessages = await getMessagesOfChannel(
      accessToken,
      teamId,
      channelId
    );
    if (config.isDebug) {
      console.log("gotmessages");
      console.log(gotmessages);
    }

    for (let i = 0, len = gotmessages.value.length; i < len; ++i) {
      var r = await getRepliesOfMessage(
        accessToken,
        teamId,
        channelId,
        gotmessages.value[i].id
      );
      if (config.isDebug) {
        console.log("gotmessages i:" + String(i) + " r.value:", r.value);
      }
      gotmessages.value[i].replies = r.value;

      this.setState({
        messages: gotmessages.value
      });
    }
    this.setState({
      selected: {
        channel: {
          id: channelId,
          name: channelName
        },
        team: {
          id: teamId,
          name: teamName
        }
      }
    });

    this.setState({
      isDropdownTreeSelectDisabled: false
    });
  }

  async ReadGraphTeamsData(force = false) {
    if (config.isDebug) {
      console.log("ReadGraphTeamsData(" + String(force) + ")");
    }

    // Get the user's access token
    var accessToken = await window.msal.acquireTokenSilent({
      scopes: config.scopes
    });

    // Get users
    if (force || !this.state.users || 0 === this.state.users.length) {
      var gotusers = await getUsers(accessToken);
      if (config.isDebug) {
        console.log("gotusers");
        console.log(gotusers);
      }

      // Update the array of users in state
      this.setState({
        users: gotusers.value
      });

      this.Notify("success", "[Graph API]ユーザー読込みが完了しました。");
    }
    if (config.isDebug) {
      console.log("this.state.users");
      console.log(this.state.users);
    }

    if (force || !this.state.teams || 0 === this.state.teams.length) {
      var gotTeams = await getJoinedTeams(accessToken);
      if (config.isDebug) {
        console.log("gotTeams.value");
        console.log(gotTeams.value);
      }
      this.setState({
        teams: gotTeams.value
      });

      this.Notify("success", "[Graph API]チーム読込みが完了しました。");
    }
    if (config.isDebug) {
      console.log("this.state.teams");
      console.log(this.state.teams);
    }

    if (force || !this.state.channels || 0 === this.state.channels.length) {
      const channels = {
        label: "Channels",
        value: "Channels",
        children: []
      };
      for (let i = 0, len = this.state.teams.length; i < len; ++i) {
        var team = this.state.teams[i]; // {};
        if (config.isDebug) {
          console.log("i: " + String(i) + " team:" + this.state.teams[i].id);
        }
        team.desc = this.state.teams[i].description;
        team.label = this.state.teams[i].displayName;
        team.value = this.state.teams[i].id;
        team.children = [];
        var gotChannels = await getChannelsOfTeam(
          accessToken,
          this.state.teams[i].id
        );
        for (let j = 0, len = gotChannels.value.length; j < len; ++j) {
          const l =
            this.state.teams[i].displayName +
            "/" +
            gotChannels.value[j].displayName;
          const v = this.state.teams[i].id + "/" + gotChannels.value[j].id;
          var channel = gotChannels.value[j]; // {};
          channel.desc = this.state.teams[i].description;
          channel.label = l;
          channel.value = v;
          team.children.push(channel);
        }
        channels.children.push(team);
        this.setState({
          channels: channels
        });
      }
    }
    if (config.isDebug) {
      console.log("this.state.channels");
      console.log(this.state.channels);
    }

    if (config.isDebug) {
      console.log("localStorage.setItem");
      console.log("JSON.stringify(this.state)");
      console.log(JSON.stringify(this.state));
      console.log("JSON.stringify(valuesToSave)");
    }
    const valuesToSave = {
      channels: this.state.channels,
      teams: this.state.teams,
      users: this.state.users
    };
    // localStorage.setItem("state", JSON.stringify(this.state));
    localStorage.setItem("state", JSON.stringify(valuesToSave));

    if (config.isDebug) {
      console.log(JSON.stringify(valuesToSave));
    }

    this.Notify("success", "[Graph API]チャネル読込みが完了しました。");
  }

  render() {
    return (
      <div>
        <link
          rel="stylesheet"
          href="https://fonts.googleapis.com/icon?family=Material+Icons"
        ></link>
        <div>
          {(() => {
            if (this.state.isAuthenticated) {
              return (
                <div>
                  ようこそ、{this.state.user.displayName} (
                  {this.state.user.email})さん{"   "}
                  <Button variant="contained" onClick={logout}>
                    サインアウト{" "}
                  </Button>
                </div>
              );
            }
          })()}{" "}
        </div>{" "}
        {(() => {
          if (this.state.channels) {
            return (
              <div>
                <div>
                  <div>
                    <DropdownTreeSelect
                      data={this.state.channels}
                      disabled={this.state.isDropdownTreeSelectDisabled}
                      mode="radioSelect"
                      onChange={this.onTreeChange}
                      onAction={this.onTreeAction}
                      onNodeToggle={this.onTreeNodeToggle}
                      texts={{ placeholder: "Select a Channel..." }}
                    />
                    {this.state.selected
                      ? (this.state.selected.team
                          ? (this.state.selected.team.name
                              ? this.state.selected.team.name
                              : "") +
                            " " +
                            (this.state.selected.team.id
                              ? "( " + this.state.selected.team.id + " )"
                              : "")
                          : "") +
                        (this.state.selected.channel
                          ? (this.state.selected.channel.name
                              ? " / " + this.state.selected.channel.name
                              : "") +
                            " " +
                            (this.state.selected.channel.id
                              ? "( " + this.state.selected.channel.id + " )"
                              : "")
                          : "")
                      : ""}{" "}
                  </div>
                  <div>
                    <input
                      type="text"
                      name="messageId"
                      placeholder="messageId"
                      value={
                        this.state.selected.message
                          ? this.state.selected.message.id
                          : ""
                      }
                      onChange={e => {
                        let s = Object.assign({}, this.state);
                        s["selected"]["message"] = {
                          id: e.target.value
                        };
                        this.setState(s);
                      }}
                    />
                    <input
                      type="text"
                      name="chatMessageText"
                      placeholder="chatMessageText"
                      value={this.state.chatMessageText}
                      onChange={e =>
                        this.setState({ chatMessageText: e.target.value })
                      }
                    />
                    <button onClick={this.postChatMessage}>Post</button>
                  </div>
                </div>
              </div>
            );
          }
        })()}{" "}
        {(() => {
          if (this.state.messages) {
            return (
              <MessageCardList messages={this.state.messages}></MessageCardList>
            );
          }
        })()}{" "}
        {(() => {
          if (this.state.messages) {
            return (
              <MaterialTable
                actions={[
                  {
                    icon: "reply",
                    tooltip: "",
                    onClick: (event, rowData) => {
                      console.log("onClick()", rowData);
                      let s = Object.assign({}, this.state);
                      s["selected"]["message"] = {
                        id: rowData.id
                      };
                      this.setState(s);
                    }
                  }
                ]}
                title="React-GraphApi-MSTeams"
                columns={this.state.columnsMessageTable}
                data={this.state.messages}
                options={{
                  pageSize: 20,
                  sorting: true
                }}
              />
            );
          }
        })()}{" "}
        {(() => {
          if (this.state.messages) {
            return <JSONTree data={this.state.messages} />;
          }
        })()}{" "}
        {(() => {
          if (this.state.channels) {
            return <JSONTree data={this.state.channels} />;
          }
        })()}{" "}
        {(() => {
          if (this.state.teams) {
            return <JSONTree data={this.state.teams} />;
          }
        })()}{" "}
        <MaterialTable
          title="React-GraphApi-MSTeams"
          columns={this.state.columnsUserTable}
          data={this.state.users}
          options={{
            pageSize: 20,
            sorting: true
          }}
        />{" "}
        <ToastContainer hideProgressBar />
      </div>
    );
  }
}
export default App;

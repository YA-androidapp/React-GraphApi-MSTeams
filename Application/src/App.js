import React, { Component } from "react";
import MaterialTable from "material-table";
// import { DragDropContext } from 'react-beautiful-dnd';
import axios from "axios";

import Button from "@material-ui/core/Button";

import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";

import {
  getChannelsOfTeam,
  getJoinedTeams,
  getMessagesOfChannel,
  getUsers
} from "./GraphService";
import { getUserDetails } from "./GraphService";
import { UserAgentApplication } from "msal";
import config from "./Config";

// yarn add react-json-tree
import JSONTree from "react-json-tree";

// yarn add react-dropdown-tree-select
import DropdownTreeSelect from "react-dropdown-tree-select";
import "react-dropdown-tree-select/dist/styles.css";

// Icons
import { forwardRef } from "react";
import AddBox from "@material-ui/icons/AddBox";
import ArrowUpward from "@material-ui/icons/ArrowUpward";
import Check from "@material-ui/icons/Check";
import ChevronLeft from "@material-ui/icons/ChevronLeft";
import ChevronRight from "@material-ui/icons/ChevronRight";
import Clear from "@material-ui/icons/Clear";
import DeleteOutline from "@material-ui/icons/DeleteOutline";
import Edit from "@material-ui/icons/Edit";
import FilterList from "@material-ui/icons/FilterList";
import FirstPage from "@material-ui/icons/FirstPage";
import LastPage from "@material-ui/icons/LastPage";
import Remove from "@material-ui/icons/Remove";
import SaveAlt from "@material-ui/icons/SaveAlt";
import Search from "@material-ui/icons/Search";
import ViewColumn from "@material-ui/icons/ViewColumn";
const tableIcons = {
  Add: forwardRef((props, ref) => <AddBox {...props} ref={ref} />),
  Check: forwardRef((props, ref) => <Check {...props} ref={ref} />),
  Clear: forwardRef((props, ref) => <Clear {...props} ref={ref} />),
  Delete: forwardRef((props, ref) => <DeleteOutline {...props} ref={ref} />),
  DetailPanel: forwardRef((props, ref) => (
    <ChevronRight {...props} ref={ref} />
  )),
  Edit: forwardRef((props, ref) => <Edit {...props} ref={ref} />),
  Export: forwardRef((props, ref) => <SaveAlt {...props} ref={ref} />),
  Filter: forwardRef((props, ref) => <FilterList {...props} ref={ref} />),
  FirstPage: forwardRef((props, ref) => <FirstPage {...props} ref={ref} />),
  LastPage: forwardRef((props, ref) => <LastPage {...props} ref={ref} />),
  NextPage: forwardRef((props, ref) => <ChevronRight {...props} ref={ref} />),
  PreviousPage: forwardRef((props, ref) => (
    <ChevronLeft {...props} ref={ref} />
  )),
  ResetSearch: forwardRef((props, ref) => <Clear {...props} ref={ref} />),
  Search: forwardRef((props, ref) => <Search {...props} ref={ref} />),
  SortArrow: forwardRef((props, ref) => <ArrowUpward {...props} ref={ref} />),
  ThirdStateCheck: forwardRef((props, ref) => <Remove {...props} ref={ref} />),
  ViewColumn: forwardRef((props, ref) => <ViewColumn {...props} ref={ref} />)
};
//

const FUNCTIONS_BASEURI = "https://example.azurewebsites.net/api/v1/User";
const FUNCTIONS_KEY = "?code=********************";

class App extends Component {
  constructor(props) {
    super(props);

    if (config.isDebug) {
      console.log("isDebug");
    }

    this.userAgentApplication = new UserAgentApplication({
      auth: {
        clientId: config.appId
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
      }
    });
    if (config.isDebug) {
      console.log("this.userAgentApplication");
      console.log(this.userAgentApplication);
    }

    var user = this.userAgentApplication.getAccount();
    if (config.isDebug) {
      console.log("user");
      console.log(user);
    }

    this.state = {
      isAuthenticated: user !== null,

      columns: [
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

      channels: [],
      teams: [],
      user: {},
      users: []
    };

    if (config.isDebug) {
      console.log("this.state");
      console.log(this.state);
    }

    if (user) {
      // Enhance user object with data from Graph
      console.log("if (user)");
      this.getUserProfile();
    } else {
      console.log("if (!user)");
      this.login();
    }

    this.ReadJoinGraphUsers = this.ReadJoinGraphUsers.bind(this);
    this.signout = this.signout.bind(this);
    this.Notify = this.Notify.bind(this);
    this.OnClickUpdate = this.OnClickUpdate.bind(this);
    this.OnClickRemove = this.OnClickRemove.bind(this);
    this.onTreeChange = this.onTreeChange.bind(this);
    this.onTreeAction = this.onTreeAction.bind(this);
    this.onTreeNodeToggle = this.onTreeNodeToggle.bind(this);

    console.log("this.ReadJoinGraphUsers()");
    this.ReadJoinGraphUsers();
  }

  async onTreeChange(currentNode, selectedNodes) {
    console.log("onTreeChange:", currentNode, selectedNodes);

    // Get the user's accessr token
    var accessToken = await window.msal.acquireTokenSilent({
      scopes: config.scopes
    });
    if (config.isDebug) {
      console.log("accessToken");
      console.log(accessToken);
    }

    console.log("表示するチャネル:", currentNode.label, currentNode.value);
    var s = currentNode.value.split(" ");
    var gotmessages = await getMessagesOfChannel(accessToken, s[0], s[1]);
    this.setState({
      messages: gotmessages.value
    });
  }

  onTreeAction = (node, action) => {
    console.log("onTreeAction:", action, node);
  };

  onTreeNodeToggle = currentNode => {
    console.log("onTreeNodeToggle:", currentNode);
  };

  signout = () => {
    this.userAgentApplication.logout();
  };

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

  async ReadJoinGraphUsers() {
    if (config.isDebug) {
      console.log("ReadJoinGraphUsers()");
    }

    try {
      console.log("ReadJoinGraphUsers");
      // Get the user's accessr token
      var accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });
      if (config.isDebug) {
        console.log("accessToken");
        console.log(accessToken);
      }
      // Get users
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
      if (config.isDebug) {
        console.log("gotusers.value");
        console.log(gotusers.value);
      }

      var gotTeams = await getJoinedTeams(accessToken);
      if (config.isDebug) {
        console.log("gotTeams.value");
        console.log(gotTeams.value);
      }
      this.setState({
        teams: gotTeams.value
      });

      this.Notify("success", "[Graph API]チーム読込みが完了しました。");

      if (config.isDebug) {
        console.log("allChannels = {}");
      }
      const allChannels = {
        label: "Channels",
        value: "Channels",
        children: []
      };
      for (let i = 0, len = gotTeams.value.length; i < len; ++i) {
        var team = {}; // gotTeams.value[i];
        console.log("i: " + String(i) + " team:" + gotTeams.value[i].id);
        team.label = gotTeams.value[i].displayName;
        team.value = gotTeams.value[i].id;
        team.children = [];
        var gotChannels = await getChannelsOfTeam(
          accessToken,
          gotTeams.value[i].id
        );
        for (let j = 0, len = gotChannels.value.length; j < len; ++j) {
          var channel = {};
          channel.label = gotChannels.value[j].displayName;
          channel.value = gotTeams.value[i].id + " " + gotChannels.value[j].id;
          team.children.push(channel);
        }
        allChannels.children.push(team);
      }
      if (config.isDebug) {
        console.log("allChannels");
        console.log(allChannels);
      }
      this.setState({
        channels: allChannels
      });
      if (config.isDebug) {
        console.log("this.state.channels");
        console.log(this.state.channels);
      }

      this.Notify("success", "[Graph API]チャネル読込みが完了しました。");
    } catch (err) {
      this.Notify(
        "error",
        "エラーが発生しました: " +
          err.message +
          " : " +
          err.fileName +
          ":" +
          err.lineNumber
      );
    }
  }

  OnClickUpdate(aadid) {
    if (config.isDebug) {
      console.log("OnClickUpdate(" + String(aadid) + ")");
    }
  }

  OnClickRemove(aadid) {
    if (config.isDebug) {
      console.log("OnClickRemove(" + String(aadid) + ")");
    }

    axios
      .delete(FUNCTIONS_BASEURI + "/" + aadid + FUNCTIONS_KEY)
      .then(results => {
        const status = results.status;
        if (config.isDebug) {
          console.log("status");
          console.log(status);
        }
        if (status.toString() === "204") {
          this.Notify("success", "[FUNCTIONS]削除が完了しました。");
        }
      })
      .catch(err => {
        this.Notify(
          "error",
          "エラーが発生しました: " +
            err.message +
            " : " +
            err.fileName +
            ":" +
            err.lineNumber
        );
      });
  }

  async login() {
    if (config.isDebug) {
      console.log("login()");
    }

    try {
      await this.userAgentApplication.loginPopup({
        scopes: config.scopes,
        prompt: "select_account"
      });
      await this.getUserProfile();
    } catch (err) {
      this.setState({
        isAuthenticated: false,
        user: {}
      });
      this.Notify(
        "error",
        "エラーが発生しました: " +
          err.message +
          " : " +
          err.fileName +
          ":" +
          err.lineNumber
      );
    }
  }

  async logout() {
    if (config.isDebug) {
      console.log("logout()");
    }

    await this.userAgentApplication.logout();
  }

  async getUserProfile() {
    if (config.isDebug) {
      console.log("getUserProfile()");
    }

    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });
      if (config.isDebug) {
        console.log("accessToken");
        console.log(accessToken);
      }

      if (accessToken) {
        // Get the user's profile from Graph
        var usr = await getUserDetails(accessToken);
        console.log("usr");
        console.log(usr);
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: usr.displayName,
            email: usr.mail || usr.userPrincipalName
          }
        });
      }
    } catch (err) {
      this.setState({
        isAuthenticated: false,
        user: null
      });
      this.Notify(
        "error",
        "エラーが発生しました: " +
          err.message +
          " : " +
          err.fileName +
          ":" +
          err.lineNumber
      );
    }
  }

  render() {
    if (config.isDebug) {
      console.log("render()");
    }

    return (
      <div>
        <div>
          {" "}
          {(() => {
            if (this.state.isAuthenticated) {
              return (
                <div>
                  <div>
                    ようこそ、{this.state.user.displayName} (
                    {this.state.user.email})さん
                  </div>
                  <Button variant="contained" onClick={this.signout}>
                    サインアウト{" "}
                  </Button>
                </div>
              );
            }
          })()}{" "}
        </div>{" "}
        {(() => {
          if (this.state.messages) {
            return <JSONTree data={this.state.messages} />;
          }
        })()}{" "}
        {(() => {
          if (this.state.channels) {
            return (
              <DropdownTreeSelect
                data={this.state.channels}
                mode="radioSelect"
                onChange={this.onTreeChange}
                onAction={this.onTreeAction}
                onNodeToggle={this.onTreeNodeToggle}
              />
            );
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
          icons={tableIcons}
          title="React-GraphApi-MSTeams"
          columns={this.state.columns}
          data={this.state.users}
          editable={{
            onRowAdd: newData =>
              new Promise((resolve, reject) => {
                setTimeout(() => {
                  axios
                    .post(FUNCTIONS_BASEURI + FUNCTIONS_KEY, newData)
                    .then(results => {
                      const status = results.status;
                      if (config.isDebug) {
                        console.log("status");
                        console.log(status);
                      }
                      if (status.toString() === "201") {
                        this.Notify(
                          "success",
                          "[FUNCTIONS]追加が完了しました。"
                        );

                        const data = this.state.users;
                        if (config.isDebug) {
                          console.log("data");
                          console.log(data);
                        }
                        data.push(newData);
                        this.setState(
                          {
                            data
                          },
                          () => resolve()
                        );
                      }
                    })
                    .catch(() => {
                      console.log("[FUNCTIONS]通信に失敗しました。");
                    });
                  resolve();
                }, 1000);
              }),
            onRowUpdate: (newData, oldData) =>
              new Promise((resolve, reject) => {
                setTimeout(() => {
                  axios
                    .put(
                      FUNCTIONS_BASEURI + "/" + oldData.aadid + FUNCTIONS_KEY,
                      newData
                    )
                    .then(results => {
                      const status = results.status;
                      if (config.isDebug) {
                        console.log("status");
                        console.log(status);
                      }
                      if (status.toString() === "204") {
                        this.Notify(
                          "success",
                          "[FUNCTIONS]保存が完了しました。"
                        );

                        const data = this.state.users;
                        if (config.isDebug) {
                          console.log("data");
                          console.log(data);
                        }
                        const index = data.indexOf(oldData);
                        if (config.isDebug) {
                          console.log("index");
                          console.log(index);
                        }
                        data[index] = newData;
                        this.setState(
                          {
                            data
                          },
                          () => resolve()
                        );
                      }
                    })
                    .catch(err => {
                      this.Notify(
                        "error",
                        "エラーが発生しました: " +
                          err.message +
                          " : " +
                          err.fileName +
                          ":" +
                          err.lineNumber
                      );
                    });
                  resolve();
                }, 1000);
              }),
            onRowDelete: oldData =>
              new Promise((resolve, reject) => {
                setTimeout(() => {
                  axios
                    .delete(
                      FUNCTIONS_BASEURI + "/" + oldData.aadid + FUNCTIONS_KEY
                    )
                    .then(results => {
                      const status = results.status;
                      if (config.isDebug) {
                        console.log("status");
                        console.log(status);
                      }
                      if (status.toString() === "204") {
                        this.Notify(
                          "success",
                          "[FUNCTIONS]削除が完了しました。"
                        );

                        let data = this.state.users;
                        if (config.isDebug) {
                          console.log("data");
                          console.log(data);
                        }
                        const index = data.indexOf(oldData);
                        if (config.isDebug) {
                          console.log("index");
                          console.log(index);
                        }
                        data.splice(index, 1);
                        this.setState(
                          {
                            data
                          },
                          () => resolve()
                        );
                      }
                    })
                    .catch(err => {
                      this.Notify(
                        "error",
                        "エラーが発生しました: " +
                          err.message +
                          " : " +
                          err.fileName +
                          ":" +
                          err.lineNumber
                      );
                    });
                  resolve();
                }, 1000);
              })
          }}
          options={{
            pageSize: 10,
            sorting: true
          }}
        />{" "}
        <ToastContainer hideProgressBar />
      </div>
    );
  }
}
export default App;

import React, { Component } from "react";
import MaterialTable from "material-table";
// import { DragDropContext } from 'react-beautiful-dnd';
import axios from 'axios';

import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';

import { getUsers } from './GraphService';
import { getUserDetails } from './GraphService';
import { UserAgentApplication } from 'msal';
import config from './Config';

// Icons
import { forwardRef } from 'react';
import AddBox from '@material-ui/icons/AddBox';
import ArrowUpward from '@material-ui/icons/ArrowUpward';
import Check from '@material-ui/icons/Check';
import ChevronLeft from '@material-ui/icons/ChevronLeft';
import ChevronRight from '@material-ui/icons/ChevronRight';
import Clear from '@material-ui/icons/Clear';
import DeleteOutline from '@material-ui/icons/DeleteOutline';
import Edit from '@material-ui/icons/Edit';
import FilterList from '@material-ui/icons/FilterList';
import FirstPage from '@material-ui/icons/FirstPage';
import LastPage from '@material-ui/icons/LastPage';
import Remove from '@material-ui/icons/Remove';
import SaveAlt from '@material-ui/icons/SaveAlt';
import Search from '@material-ui/icons/Search';
import ViewColumn from '@material-ui/icons/ViewColumn';
const tableIcons = {
  Add: forwardRef((props, ref) => <AddBox {...props} ref={ref} />),
  Check: forwardRef((props, ref) => <Check {...props} ref={ref} />),
  Clear: forwardRef((props, ref) => <Clear {...props} ref={ref} />),
  Delete: forwardRef((props, ref) => <DeleteOutline {...props} ref={ref} />),
  DetailPanel: forwardRef((props, ref) => <ChevronRight {...props} ref={ref} />),
  Edit: forwardRef((props, ref) => <Edit {...props} ref={ref} />),
  Export: forwardRef((props, ref) => <SaveAlt {...props} ref={ref} />),
  Filter: forwardRef((props, ref) => <FilterList {...props} ref={ref} />),
  FirstPage: forwardRef((props, ref) => <FirstPage {...props} ref={ref} />),
  LastPage: forwardRef((props, ref) => <LastPage {...props} ref={ref} />),
  NextPage: forwardRef((props, ref) => <ChevronRight {...props} ref={ref} />),
  PreviousPage: forwardRef((props, ref) => <ChevronLeft {...props} ref={ref} />),
  ResetSearch: forwardRef((props, ref) => <Clear {...props} ref={ref} />),
  Search: forwardRef((props, ref) => <Search {...props} ref={ref} />),
  SortArrow: forwardRef((props, ref) => <ArrowUpward {...props} ref={ref} />),
  ThirdStateCheck: forwardRef((props, ref) => <Remove {...props} ref={ref} />),
  ViewColumn: forwardRef((props, ref) => <ViewColumn {...props} ref={ref} />),
};
//


const FUNCTIONS_BASEURI = 'https://example.azurewebsites.net/api/v1/User';
const FUNCTIONS_KEY = '?code=********************';

class App extends Component {
  constructor(props) {
    super(props);

    if (config.isDebug) { console.log('isDebug'); }

    this.userAgentApplication = new UserAgentApplication({
      auth: {
        clientId: config.appId
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true
      }
    });
    if (config.isDebug) { console.log('this.userAgentApplication'); console.log(this.userAgentApplication); }

    var user = this.userAgentApplication.getAccount();
    if (config.isDebug) { console.log('user'); console.log(user); }

    this.state = {
      isAuthenticated: (user !== null),

      columns: [
        { title: 'ID', field: 'id', editable: 'never' },
        { title: 'DisplayName', field: 'displayName' },
        { title: 'userPrincipalName', field: 'userPrincipalName' },
        { title: 'mobilePhone', field: 'mobilePhone' }
      ],

      user: {},
      users: [],
      dbusers: [],
      graphusers: [],
    };
    if (config.isDebug) { console.log('this.state'); console.log(this.state); }

    if (user) {
      // Enhance user object with data from Graph
      console.log('if (user)');
      this.getUserProfile();
    } else {
      console.log('if (!user)');
      this.login();
    }

    this.ReadUsers = this.ReadUsers.bind(this);
    this.Notify = this.Notify.bind(this);
    this.OnClickUpdate = this.OnClickUpdate.bind(this);
    this.OnClickRemove = this.OnClickRemove.bind(this);

    // this.ReadUsers();
    this.JoinGraphUsers();
  }

  Notify = (type, message) => {
    if (config.isDebug) { console.log('Notify(' + String(type) + ', ' + String(message) + ')'); }

    var date = new Date();
    if (config.isDebug) { console.log('date'); console.log(date); }
    var toastId = date.getTime(); // UNIXTIME(msec)
    if (config.isDebug) { console.log('toastId'); console.log(toastId); }

    switch (type) {
      case "info":
        toast.info(message, { toastId: toastId });
        break;
      case "success":
        toast.success(message, { toastId: toastId });
        break;
      case "warning":
        toast.warn(message, { toastId: toastId });
        break;
      case "error":
        toast.error(message, { toastId: toastId });
        break;
      default:
        toast(message, { toastId: toastId });
        break;
    }
  }

  ReadUsers() {
    if (config.isDebug) { console.log('ReadUsers()'); }

  //   axios
  //     .get(FUNCTIONS_BASEURI + FUNCTIONS_KEY)
  //     .then((results) => {
  //       const status = results.status;
  //       if (config.isDebug) { console.log('status'); console.log(status); }
  //       if (status.toString() === "200") {
  //         const data = results.data;
  //         if (config.isDebug) { console.log('data'); console.log(data); }
  //         this.setState({ dbusers: data });
  //         this.Notify("info", "[FUNCTIONS]読込みが完了しました。");
  //       }
  //     },
  //     )
  //     .catch((e) => {
  //     });
  }

  async JoinGraphUsers() {
    if (config.isDebug) { console.log('JoinGraphUsers()'); }

    try {
      console.log('JoinGraphUsers')
      // Get the user's accessr token
      var accessToken = await window.msal.acquireTokenSilent({
        scopes: config.scopes
      });
      console.log(accessToken)
      // Get users
      var gotusers = await getUsers(accessToken);
      console.log('gotusers'); console.log(gotusers);
      // Update the array of users in state
      this.setState({ graphusers: gotusers.value });

      this.Notify("success", "[Graph API]読込みが完了しました。");
      console.log('gotusers.value');
      console.log(gotusers.value);

      this.setState({ users: this.state.graphusers });

      // join
      // var DbUsers = this.state.dbusers;
      // var GraphUsers = this.state.graphusers;
      // console.log('DbUsers');
      // console.log(DbUsers);
      // console.log('GraphUsers');
      // console.log(GraphUsers);
      //
      // if ((typeof DbUsers !== 'undefined') && (typeof GraphUsers !== 'undefined')) {
      //   if ((DbUsers.length > 0) && (GraphUsers.length > 0)) {
      //     console.log('DbUsers');
      //     console.log(DbUsers);
      //     console.log('GraphUsers');
      //     console.log(GraphUsers);
      //     var Joined = DbUsers.map(dbusr =>
      //       GraphUsers.some(gusr => gusr.id === dbusr.aadid) ?
      //         GraphUsers.filter(gusr => gusr.id === dbusr.aadid).map(gusr => Object.assign(gusr, dbusr)) :
      //         { dbusr }
      //     ).reduce((a, b) => a.concat(b), []);
      //     console.log('Joined');
      //     console.log(Joined);
      //
      //     this.setState({ users: Joined });
      //   }
      // }
      //
    }
    catch (err) {
      console.log(String(err));
    }
  }

  OnClickUpdate(aadid) {
    if (config.isDebug) { console.log('OnClickUpdate(' + String(aadid) + ')'); }
  }

  OnClickRemove(aadid) {
    if (config.isDebug) { console.log('OnClickRemove(' + String(aadid) + ')'); }

    axios
      .delete(FUNCTIONS_BASEURI + '/' + aadid + FUNCTIONS_KEY)
      .then((results) => {
        const status = results.status;
        if (config.isDebug) { console.log('status'); console.log(status); }
        if (status.toString() === "204") {
          this.Notify("success", "[FUNCTIONS]削除が完了しました。");
        }
      },
      )
      .catch((e) => {
        this.Notify("error", "[FUNCTIONS]通信に失敗しました。" + e);
      });
  }

  async login() {
    if (config.isDebug) { console.log('login()'); }

    try {
      await this.userAgentApplication.loginPopup(
        {
          scopes: config.scopes,
          prompt: "select_account"
        });
      await this.getUserProfile();
    }
    catch (err) {
      if (typeof (err) === 'string') {
        var errParts = err.split('|');
        this.setState({
          isAuthenticated: false,
          user: {},
          error: { message: errParts[1], debug: errParts[0] }
        });
      } else {
        this.setState({
          isAuthenticated: false,
          user: {},
          error: err
        });
      }
    }
  }

  logout() {
    if (config.isDebug) { console.log('logout()'); }

    this.userAgentApplication.logout();
  }

  async getUserProfile() {
    if (config.isDebug) { console.log('getUserProfile()'); }

    try {
      // Get the access token silently
      // If the cache contains a non-expired token, this function
      // will just return the cached token. Otherwise, it will
      // make a request to the Azure OAuth endpoint to get a token

      var accessToken = await this.userAgentApplication.acquireTokenSilent({
        scopes: config.scopes
      });
      console.log('accessToken'); console.log(accessToken);

      if (accessToken) {
        // Get the user's profile from Graph
        var usr = await getUserDetails(accessToken);
        console.log('usr'); console.log(usr);
        this.setState({
          isAuthenticated: true,
          user: {
            displayName: usr.displayName,
            email: usr.mail || usr.userPrincipalName
          },
          error: null
        });
      }
    }
    catch (err) {
      console.log('err'); console.log(err);

      var error = {};
      if (typeof (err) === 'string') {
        var errParts = err.split('|');
        error = errParts.length > 1 ?
          { message: errParts[1], debug: errParts[0] } :
          { message: err };
      } else {
        error = {
          message: err.message,
          debug: JSON.stringify(err)
        };
      }

      this.setState({
        isAuthenticated: false,
        user: null,
        error: error
      });
    }
  }

  render() {
    if (config.isDebug) { console.log('render()'); }

    return (
      <div>
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
                    .then((results) => {
                      const status = results.status;
                      if (config.isDebug) { console.log('status'); console.log(status); }
                      if (status.toString() === "201") {
                        this.Notify("success", "[FUNCTIONS]追加が完了しました。");

                        const data = this.state.users;
                        if (config.isDebug) { console.log('data'); console.log(data); }
                        data.push(newData);
                        this.setState({ data }, () => resolve());
                      }
                    },
                    )
                    .catch(() => {
                      console.log('[FUNCTIONS]通信に失敗しました。');
                    });
                  resolve()
                }, 1000)
              }),
            onRowUpdate: (newData, oldData) =>
              new Promise((resolve, reject) => {
                setTimeout(() => {
                  axios
                    .put(FUNCTIONS_BASEURI + '/' + oldData.aadid + FUNCTIONS_KEY, newData)
                    .then((results) => {
                      const status = results.status;
                      if (config.isDebug) { console.log('status'); console.log(status); }
                      if (status.toString() === "204") {
                        this.Notify("success", "[FUNCTIONS]保存が完了しました。");

                        const data = this.state.users;
                        if (config.isDebug) { console.log('data'); console.log(data); }
                        const index = data.indexOf(oldData);
                        if (config.isDebug) { console.log('index'); console.log(index); }
                        data[index] = newData;
                        this.setState({ data }, () => resolve());
                      }
                    },
                    )
                    .catch(() => {
                      this.Notify("error", "[FUNCTIONS]通信に失敗しました。");
                    });
                  resolve()
                }, 1000)
              }),
            onRowDelete: oldData =>
              new Promise((resolve, reject) => {
                setTimeout(() => {
                  axios
                    .delete(FUNCTIONS_BASEURI + '/' + oldData.aadid + FUNCTIONS_KEY)
                    .then((results) => {
                      const status = results.status;
                      if (config.isDebug) { console.log('status'); console.log(status); }
                      if (status.toString() === "204") {
                        this.Notify("success", "[FUNCTIONS]削除が完了しました。");

                        let data = this.state.users;
                        if (config.isDebug) { console.log('data'); console.log(data); }
                        const index = data.indexOf(oldData);
                        if (config.isDebug) { console.log('index'); console.log(index); }
                        data.splice(index, 1);
                        this.setState({ data }, () => resolve());
                      }
                    },
                    )
                    .catch((e) => {
                      this.Notify("error", "[FUNCTIONS]通信に失敗しました。" + e);
                    });
                  resolve()
                }, 1000)
              }),
          }}
          options={{
            pageSize: 10,
            sorting: true
          }}
        />
        <ToastContainer hideProgressBar />
      </div>
    )
  }
}
export default App;

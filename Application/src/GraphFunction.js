import { UserAgentApplication } from "msal";
import config from "./Config";
import { getUserDetails } from "./GraphService";

export async function setupUserAgentApplication(app) {
  let userAgentApplication = new UserAgentApplication({
    auth: {
      clientId: config.appId
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  });
  if (config.isDebug) {
    console.log("userAgentApplication");
    console.log(userAgentApplication);
  }

  var user = await userAgentApplication.getAccount();
  if (config.isDebug) {
    console.log("user");
    console.log(user);
  }

  if (user) {
    // Enhance user object with data from Graph
    console.log("if (user)");
    await getUserProfile(app, userAgentApplication);
    return userAgentApplication;
  } else {
    console.log("if (!user)");
    await login(app, userAgentApplication);
    var userRetry = await userAgentApplication.getAccount();
    if (config.isDebug) {
      console.log("if (!user) login() userRetry");
      console.log(userRetry);
    }
    return userAgentApplication;
  }
}

export async function login(app, userAgentApplication) {
  if (config.isDebug) {
    console.log("login()");
  }

  await userAgentApplication.loginPopup({
    scopes: config.scopes,
    prompt: "select_account"
  });
  await getUserProfile(app, userAgentApplication);
}

export async function logout(userAgentApplication) {
  if (config.isDebug) {
    console.log("logout()");
  }

  await userAgentApplication.logout();
}

export async function getUserProfile(app, userAgentApplication) {
  if (config.isDebug) {
    console.log("getUserProfile(app, userAgentApplication)");
  }

  var accessToken = await userAgentApplication.acquireTokenSilent({
    scopes: config.scopes
  });

  if (accessToken) {
    // Get the user's profile from Graph
    var usr = await getUserDetails(accessToken);
    console.log("usr");
    console.log(usr);
    app.setState({
      isAuthenticated: true,
      user: {
        displayName: usr.displayName,
        email: usr.mail || usr.userPrincipalName
      }
    });
    console.log("setState");
  }
}

var graph = require('@microsoft/microsoft-graph-client');

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken.accessToken);
    }
  });

  return client;
}

export async function getUserDetails(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const user = await client.api('/me').get();
  return user;
}

export async function getJoinedTeams(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const teams = await client
    .api('/me/joinedTeams')
    .select('id,displayName,description,internalId,classification,specialization,visibility,webUrl,isArchived,memberSettings,guestSettings,messagingSettings,funSettings,discoverySettings')
    .orderby('displayName')
    // .orderby('internalId DESC')
    .get();

  return teams;
}

export async function getUsers(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const users = await client
    .api('/users')
    //.select('businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName')
    .get();

  console.log(users)
  return users;
}
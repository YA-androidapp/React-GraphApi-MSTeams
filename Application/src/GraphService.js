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
    .get();
  // console.log('teams');
  // console.log(teams);
  return teams;
}


export async function getMembersOfTeam(accessToken, teamId) {
  const client = getAuthenticatedClient(accessToken);

  const members = await client
    .api('/groups/' + teamId + '/members')
    .get();
  // console.log('members');
  // console.log(members);
  return members;
}


export async function getChannelsOfTeam(accessToken, teamId) {
  const client = getAuthenticatedClient(accessToken);

  const channels = await client
    .api('/teams/' + teamId + '/channels')
    .get();
  // console.log('channels');
  // console.log(channels);
  return channels;
}

export async function getMessagesOfChannel(accessToken, teamId, channelId) {
  const client = getAuthenticatedClient(accessToken);

  const messages = await client
    .api('https://graph.microsoft.com/beta/teams/' + teamId + '/channels/' + channelId + '/messages')
    .get();
  // console.log('messages');
  // console.log(messages);
  return messages;
}

export async function getRepliesOfMessage(accessToken, teamId, channelId, messageId) {
  const client = getAuthenticatedClient(accessToken);

  const messages = await client
    .api('https://graph.microsoft.com/beta/teams/' + teamId + '/channels/' + channelId + '/messages/' + messageId + '/replies')
    .get();
  // console.log('messages');
  // console.log(messages);
  return messages;
}


export async function getUsers(accessToken) {
  const client = getAuthenticatedClient(accessToken);

  const users = await client
    .api('/users')
    //.select('businessPhones,displayName,givenName,id,jobTitle,mail,mobilePhone,officeLocation,preferredLanguage,surname,userPrincipalName')
    // .orderby('internalId DESC')
    .get();

  // console.log(users)
  return users;
}
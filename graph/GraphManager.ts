// Copyright (c) Microsoft.
// Licensed under the MIT license.

import {Client} from '@microsoft/microsoft-graph-client';
import {GraphAuthProvider} from './GraphAuthProvider';

// Set the authProvider to an instance
// of GraphAuthProvider
const clientOptions = {
  authProvider: new GraphAuthProvider(),
};

// Initialize the client
const graphClient = Client.initWithMiddleware(clientOptions);

export class GraphManager {
  static getUserAsync = async () => {
    // GET /me
    return await graphClient
      .api('/me')
      .select('displayName,givenName,mail,mailboxSettings,userPrincipalName')
      .get();
  };

  // <GetCalendarViewSnippet>
  static getCalendarView = async (
    start: string,
    end: string,
    timezone: string,
  ) => {
    // GET /me/calendarview
    return await graphClient
      .api('/me/calendarview')
      .header('Prefer', `outlook.timezone="${timezone}"`)
      .query({startDateTime: start, endDateTime: end})
      // $select='subject,organizer,start,end'
      // Only return these fields in results
      .select('subject,organizer,start,end')
      // $orderby=createdDateTime DESC
      // Sort results by when they were created, newest first
      .orderby('start/dateTime')
      .top(50)
      .get();
  };
  // </GetCalendarViewSnippet>

  // <CreateEventSnippet>
  static createEvent = async (newEvent: any) => {
    // POST /me/events
    const onlineMeeting = {
      startDateTime: newEvent?.start?.dateTime,
      endDateTime: newEvent?.end?.dateTime,
      subject: newEvent?.subject,
    };

    console.log('onlineMeeting: ', onlineMeeting);

    const response = await graphClient
      .api('/me/onlineMeetings')
      .post(onlineMeeting);

    console.log('onlineMeetings response: ', response, newEvent);
    let event = newEvent;

    if (response) {
      const content = newEvent?.body?.content + '###' + response?.joinWebUrl;

      event.body.content = content;
      console.log('event: ', event);

      await graphClient
        .api('/me/events')
        .post(event)
        .then(r => {
          console.log('response event: ', r);
        });
    }
  };

  static createMeeting = async (newMeeting: any, response: any) => {
    const userToken = await clientOptions.authProvider.getAccessToken();

    console.log('createMeeting', newMeeting, userToken);

    const onlineMeeting = {
      startDateTime: '2019-07-12T14:30:34.2444915-07:00',
      endDateTime: '2019-07-12T15:00:34.2464912-07:00',
      subject: 'Meeting',
    };

    await graphClient
      .api('/me/onlineMeetings')
      .post(onlineMeeting)
      .then(res => {
        console.log(res);
        response(res);
      });
  };

  static getRecordCall = async (id: any, response: any) => {
    const url = `/communications/callRecords/${id}`;

    console.log('sub url: ', url);

    await graphClient.api(url).get();
    // .then(res => {
    //   console.log(res);
    //   response(res);
    // });
  };

  static createCall = async (payload: any) => {
    const call = {
      '@odata.type': '#microsoft.graph.call',
      callbackUri: 'https://contoso.com/teamsapp/api/calling',
      source: {
        '@odata.type': '#microsoft.graph.participantInfo',
        identity: {
          '@odata.type': '#microsoft.graph.identitySet',
          applicationInstance: {
            '@odata.type': '#microsoft.graph.identity',
            displayName: 'Chau',
            id: `${payload}`,
          },
        },
        countryCode: null,
        endpointType: null,
        region: null,
        languageId: null,
      },
      targets: [
        {
          '@odata.type': '#microsoft.graph.invitationParticipantInfo',
          identity: {
            '@odata.type': '#microsoft.graph.identitySet',
            phone: {
              '@odata.type': '#microsoft.graph.identity',
              id: '+84342145738',
            },
          },
        },
      ],
      requestedModalities: ['audio'],
      mediaConfig: {
        '@odata.type': '#microsoft.graph.appHostedMediaConfig',
        blob: '<Media Session Configuration>',
      },
      tenantId: 'cb49b5e3-5da4-49a2-b315-f8234af61e14',
    };
    console.log('call data: ', call);

    await graphClient
      .api('/communications/calls')
      .post(call)
      .then(r => console.log(r));
  };

  static getMe = async () => {
    let user = await graphClient.api('/me').get();
    console.log('me: ', user);
    return user;
  };

  static getEvent = async (eventId: any) => {
    let event = await graphClient
      .api(`/me/events/${eventId}`)
      .header('Prefer', 'outlook.timezone="Pacific Standard Time"')
      .select(
        'subject,body,bodyPreview,organizer,attendees,start,end,location,hideAttendees',
      )
      .get();
    let link = '';
    if (event) {
      const bodyReview = event.bodyPreview;
      link = bodyReview.split('###')[1];
      console.log('event link: ', link);
    }

    return link;
  };
}

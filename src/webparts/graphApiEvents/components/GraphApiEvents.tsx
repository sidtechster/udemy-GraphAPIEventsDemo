import * as React from 'react';
import styles from './GraphApiEvents.module.scss';
import { IGraphApiEventsProps } from './IGraphApiEventsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IGraphApiEventsState } from './IGraphApiEventsState';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class GraphApiEvents extends React.Component<IGraphApiEventsProps, IGraphApiEventsState> {
  
  constructor(props: IGraphApiEventsProps) {
    super(props);
    this.state = {
      events: []
    };
  }

  public componentDidMount(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client
        .api('/me/calendar/events')
        .version("v1.0")
        .select("*")
        .get((error: any, eventsResponse, rawResponse?: any) => {
          if(error) {
            console.error("Message is :" + error);
            return;
          }
          const calendarEvents: MicrosoftGraph.Event[] = eventsResponse.value;
          this.setState({ events: calendarEvents });
        });
    });
  }
  
  //See package-solution.json -> webApiPermissionRequests
  // Entries has to be made here
  // After deploying solution in app catalog, goto admin center > API access > Grant Permission
  
  public render(): React.ReactElement<IGraphApiEventsProps> {
    return (
      <div>
        <ul>
          {
            this.state.events.map((item, key) => 
              <li key={item.id}>
                {item.subject},{item.organizer.emailAddress.name},
                {item.start.dateTime.substr(0,10)},
                {item.start.dateTime.substr(12,5)},
                {item.end.dateTime.substr(0,10)},
                {item.end.dateTime.substr(12,5)}
              </li>
            )
          }
        </ul>
      </div>
    );
  }
}

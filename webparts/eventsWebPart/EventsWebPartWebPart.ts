import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import EventsPageWebPart from "./components/EventsPageComponent";
import { IEventsWebPartProps } from './components/IEventsWebPartProps';

export interface IEventsWebPartWebPartProps {
  description: string;
  context: WebPartContext;
}

export default class EventsWebPartWebPart extends BaseClientSideWebPart<IEventsWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEventsWebPartProps> = React.createElement(
      EventsPageWebPart,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
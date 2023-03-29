import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp/presets/all";
import * as strings from "ClientInfoWebpartWebPartStrings";
import ClientInfoWebpart from "./components/ClientInfoWebpart";
import { IClientInfoWebpartProps } from "./components/IClientInfoWebpartProps";

export interface IClientInfoWebpartWebPartProps {
  description: string;
}

export default class ClientInfoWebpartWebPart extends BaseClientSideWebPart<
  IClientInfoWebpartWebPartProps
> {

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context,
       
      });
    });
  }

  public render(): void {
    console.log('version is 1.0.1');
    const element: React.ReactElement<IClientInfoWebpartProps> = React.createElement(
      ClientInfoWebpart,
      {
        description: this.properties.description,
        spContext: this.context        
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}

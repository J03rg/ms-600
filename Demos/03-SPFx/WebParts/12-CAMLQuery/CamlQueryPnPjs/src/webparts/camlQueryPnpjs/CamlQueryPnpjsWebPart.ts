import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "CamlQueryPnpjsWebPartStrings";
import CamlQueryPnpjs from "./components/CamlQueryPnpjs";

import { setup as pnpSetup } from "@pnp/common";
import { ISPListItem } from "./SPListItem";
import { ICamlQuery } from "@pnp/sp/lists";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ICamlQueryPnpjsProps } from "../../../lib/webparts/camlQueryPnpjs/components/CamlQueryPnpjs";

export interface ICamlQueryPnpjsWebPartProps {
  description: string;
}

export default class CamlQueryPnpjsWebPart extends BaseClientSideWebPart<ICamlQueryPnpjsWebPartProps> {
  protected onInit(): Promise<void> {
    return super.onInit().then((_) => {
      pnpSetup({
        spfxContext: this.context,
      });
    });
  }

  public render(): void {
    this.execCAMLQuery().then((lis: ISPListItem[]) => {
      const element: React.ReactElement<ICamlQueryPnpjsProps> = React.createElement(
        CamlQueryPnpjs,
        {
          items: lis,
        }
      );

      ReactDom.render(element, this.domElement);
    });
  }

  public async execCAMLQuery(): Promise<ISPListItem[]> {
    const caml: ICamlQuery = {
      ViewXml:
        "<View><ViewFields><FieldRef Name='Title' /></ViewFields><RowLimit>1</RowLimit></View>",
    };
    const items = await sp.web.lists
      .getByTitle("SPRestList")
      .getItemsByCAMLQuery(caml);
    return items;
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

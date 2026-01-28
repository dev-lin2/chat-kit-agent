import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";

import ChatKitAgent from "./components/ChatKitAgent";
import { IChatKitAgentProps } from "./components/IChatKitAgentProps";
import { IChatKitAgentWebPartProps } from "./IChatKitAgentWebPartProps";

export default class ChatKitAgentWebPart extends BaseClientSideWebPart<IChatKitAgentWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IChatKitAgentProps> = React.createElement(ChatKitAgent, {
      lambdaUrl: (this.properties.lambdaUrl || "").trim(),
      workflowId: (this.properties.workflowId || "").trim(),
      greeting: (this.properties.greeting || "").trim(),
      userId: this._getUserId()
    });

    ReactDom.render(element, this.domElement);
  }

  private _getUserId(): string {
    const loginName = this.context.pageContext.user?.loginName;
    return loginName && loginName.trim() ? loginName : "sharepoint-user";
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
          header: { description: "Chat Kit Agent settings" },
          groups: [
            {
              groupName: "Configuration",
              groupFields: [
                PropertyPaneTextField("greeting", { label: "Greeting" }),
                PropertyPaneTextField("lambdaUrl", {
                  label: "Token endpoint (Lambda Function URL)",
                  placeholder: "https://xxxx.lambda-url.region.on.aws/"
                }),
                PropertyPaneTextField("workflowId", { label: "Workflow ID", placeholder: "wf_..." })
              ]
            }
          ]
        }
      ]
    };
  }
}

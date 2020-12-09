import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
import { TaskModuleRequest, TaskModuleContinueResponse } from "botbuilder";
// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/runActionMeMessageExtension/config.html")
@PreventIframe("/runActionMeMessageExtension/action.html")
export default class RunActionMeMessageExtension implements IMessagingExtensionMiddlewareProcessor {




    // handle action response in here
    // See documentation for `MessagingExtensionResult` for details
    public async onSubmitAction(context: TurnContext, value: TaskModuleRequest): Promise<MessagingExtensionResult> {

        log(value);

        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: value.data.Title
                    },
                    {
                        type: "TextBlock",
                        text: value.data.Description
                    }
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "More details",
                        data: {
                            action: "moreDetails",
                            id: "1234-5678"
                        }
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.2"
            });

        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [card]
        } as MessagingExtensionResult);
    }



}

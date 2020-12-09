import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, MessagingExtensionAttachment, ActionTypes } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
const quote = require('find-quote');

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/npmSearchMessageExtension/config.html")
export default class NpmAuthMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {

            log(context.activity.from);

            const adapter: any = context.adapter;
            const magicCode = (query.state && Number.isInteger(Number(query.state))) ? query.state : '';
            const tokenResponse = await adapter.getUserToken(context, process.env.CONNECTION_NAME, magicCode);

            if (!tokenResponse || !tokenResponse.token) {
                // There is no token, so the user has not signed in yet.
    
                // Retrieve the OAuth Sign in Link to use in the MessagingExtensionResult Suggested Actions
                const signInLink = await adapter.getSignInLink(context, process.env.CONNECTION_NAME);
                let composeExtension: MessagingExtensionResult = {
                    type: 'config',
                    suggestedActions: {
                        actions: [{
                            title: 'Sign in as user',
                            value: signInLink,
                            type: ActionTypes.OpenUrl
                        }]
                    }
                };
                return Promise.resolve(composeExtension);
            }

            const graphTokenResponse = tokenResponse.token;
            log(graphTokenResponse);

            const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: "asdasd"
                    },
                    {
                        type: "TextBlock",
                        text: "asdasdasd"
                    },
                    {
                        type: "Image",
                        url: `https://${process.env.HOSTNAME}/assets/icon.png`
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
            const preview = {
                contentType: "application/vnd.microsoft.card.thumbnail",
                content: {
                    title: "asdasdasd",
                    text: "adasdasds"
                }
            };

        if (query.parameters && query.parameters[0] && query.parameters[0].name === "initialRun") {
            // initial run

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);

        } else {
            // the rest

            if (query.parameters && query.parameters[0] && query.parameters[0].name === "Protected" && query.parameters[0].value === "signout") { 
                log("signout called");
                const adapter: any = context.adapter;
                await adapter.signOutUser(context, process.env.CONNECTION_NAME);
            }

            return Promise.resolve({
                type: "result",
                attachmentLayout: "list",
                attachments: [
                    { ...card, preview }
                ]
            } as MessagingExtensionResult);    
        }
    }

    public async onCardButtonClicked(context: TurnContext, value: any): Promise<void> {
        // Handle the Action.Submit action on the adaptive card
        if (value.action === "moreDetails") {
            log(`I got this ${value.id}`);
        }
        return Promise.resolve();
    }


    // this is used when canUpdateConfiguration is set to true
    public async onQuerySettingsUrl(context: TurnContext): Promise<{ title: string, value: string }> {
        return Promise.resolve({
            title: "NPM Search Configuration",
            value: `https://${process.env.HOSTNAME}/npmSearchMessageExtension/config.html?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`
        });
    }

    public async onSettings(context: TurnContext): Promise<void> {
        // take care of the setting returned from the dialog, with the value stored in state
        const setting = context.activity.value.state;
        log(`New setting: ${setting}`);
        return Promise.resolve();
    }

}

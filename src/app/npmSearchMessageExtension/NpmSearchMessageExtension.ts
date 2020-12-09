import * as debug from "debug";
import { PreventIframe } from "express-msteams-host";
import { TurnContext, CardFactory, MessagingExtensionQuery, MessagingExtensionResult, MessagingExtensionAttachment, AppBasedLinkQuery } from "botbuilder";
import { IMessagingExtensionMiddlewareProcessor } from "botbuilder-teams-messagingextensions";
const quote = require('find-quote');

// Initialize debug logging module
const log = debug("msteams");

@PreventIframe("/npmSearchMessageExtension/config.html")
export default class NpmSearchMessageExtension implements IMessagingExtensionMiddlewareProcessor {

    public async onQueryLink(context: TurnContext, value: AppBasedLinkQuery): Promise<MessagingExtensionResult> {

        log("link queried");
        log(value);

        const thumbnailCard = {
            "contentType": "application/vnd.microsoft.card.thumbnail",
            "content": {
              "title": "Bender",
              "subtitle": "tale of a robot who dared to love",
              "text": "Bender Bending Rodríguez is a main character in the animated television series Futurama. He was created by series creators Matt Groening and David X. Cohen, and is voiced by John DiMaggio",
              "images": [
                {
                  "url": "https://upload.wikimedia.org/wikipedia/en/a/a6/Bender_Rodriguez.png",
                  "alt": "Bender Rodríguez"
                }
              ],
              "buttons": [
                {
                  "type": "imBack",
                  "title": "Thumbs Up",
                  "image": "http://moopz.com/assets_c/2012/06/emoji-thumbs-up-150-thumb-autox125-140616.jpg",
                  "value": "I like it"
                },
                {
                  "type": "imBack",
                  "title": "Thumbs Down",
                  "image": "http://yourfaceisstupid.com/wp-content/uploads/2014/08/thumbs-down.png",
                  "value": "I don't like it"
                },
                {
                  "type": "openUrl",
                  "title": "I feel lucky",
                  "image": "http://thumb9.shutterstock.com/photos/thumb_large/683806/148441982.jpg",
                  "value": "https://www.bing.com/images/search?q=bender&qpvt=bender&qpvt=bender&qpvt=bender&FORM=IGRE"
                }
              ],
              "tap": {
                "type": "imBack",
                "value": "Tapped it!"
              }
            }
          }
        
        return Promise.resolve({
            type: "result",
            attachmentLayout: "list",
            attachments: [
                { ...thumbnailCard }
            ]
        } as MessagingExtensionResult);
    }

    public async onQuery(context: TurnContext, query: MessagingExtensionQuery): Promise<MessagingExtensionResult> {

        const myQuote = await quote.getQuoteWithAuthor();

        const card = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Large",
                        text: myQuote.author
                    },
                    {
                        type: "TextBlock",
                        text: myQuote.quote
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
                title: myQuote.author,
                text: myQuote.quote
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
            if (query.parameters && query.parameters[0] && query.parameters[0].name === "NPM") {

                log("success");

                const queryValue = query.parameters[0].value;
                const attachments = this.CreateAttachment(queryValue);
                
                return Promise.resolve({
                    type: "result",
                    attachmentLayout: "list",
                    attachments: attachments
                } as MessagingExtensionResult);

            } else  {

                log("error");
                return Promise.resolve({
                    type: "result",
                    attachmentLayout: "list",
                    attachments: [
                        { ...card, preview }
                    ]
                } as MessagingExtensionResult);    
            }
        }
    }

    public CreateAttachment = (query: string) => {
        
        const mySearchedQuote = quote.getQuoteWithAuthor(query);
        const attachments: MessagingExtensionAttachment[] = [];
        
        const queryCard = CardFactory.adaptiveCard(
            {
                type: "AdaptiveCard",
                body: [
                    {
                        type: "TextBlock",
                        size: "Medium",
                        weight: "Bolder",
                        text: mySearchedQuote.author
                    },
                    {
                        type: "TextBlock",
                        text: mySearchedQuote.quote,
                        wrap: true
                    }
                ],
                actions: [
                    {
                        type: "Action.OpenUrl",
                        title: "Feedback",
                        url: "https://www.forms.com"
                    }
                ],
                $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                version: "1.2"
            });

        const queryPreview = {
            contentType: "application/vnd.microsoft.card.thumbnail",
            content: {
                title: mySearchedQuote.author,
                text: mySearchedQuote.quote
            }
        };

        attachments.push({ contentType: queryCard.contentType, content: queryCard.content, preview: queryPreview });
        return attachments;
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

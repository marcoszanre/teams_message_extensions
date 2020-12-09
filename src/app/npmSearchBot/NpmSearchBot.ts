import { BotDeclaration, MessageExtensionDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import NpmSearchMessageExtension from "../npmSearchMessageExtension/NpmSearchMessageExtension";
import NpmAuthMessageExtension from "../npmSearchMessageExtension/NpmAuthMessageExtension";
import RunActionMeMessageExtension from "../runActionMeMessageExtension/RunActionMeMessageExtension";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, BotFrameworkAdapter, AppBasedLinkQuery, MessagingExtensionResult } from "botbuilder";

// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for NPM Search Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class NpmSearchBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    /** Local property for RunActionMeMessageExtension */
    @MessageExtensionDeclaration("runActionMeMessageExtension")
    private _runActionMeMessageExtension: RunActionMeMessageExtension;
    /** Local property for NpmSearchMessageExtension */
    @MessageExtensionDeclaration("npmSearchMessageExtension")
    private _npmSearchMessageExtension: NpmSearchMessageExtension;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    @MessageExtensionDeclaration("npmAuthMessageExtension")
    private _npmAuthMessageExtension: NpmAuthMessageExtension;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();
        // Message extension RunActionMeMessageExtension
        this._runActionMeMessageExtension = new RunActionMeMessageExtension();

        // Message extension NpmSearchMessageExtension
        this._npmSearchMessageExtension = new NpmSearchMessageExtension();
        this._npmAuthMessageExtension = new NpmAuthMessageExtension();


        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
    }


}

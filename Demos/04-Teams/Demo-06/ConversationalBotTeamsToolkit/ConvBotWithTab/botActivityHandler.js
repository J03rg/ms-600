// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { TurnContext, MessageFactory, TeamsActivityHandler, CardFactory, ActionTypes } = require("botbuilder");

class BotActivityHandler extends TeamsActivityHandler {
  constructor() {
    super();
    // Registers an activity event handler for the message event, emitted for every incoming message activity.
    this.onMessage(async (context, next) => {
      TurnContext.removeRecipientMention(context.activity);
      switch (context.activity.text.trim()) {
        case "Hello":
          await this.mentionActivityAsync(context);
          break;
        default:
          // By default for unknown activity sent by user show
          // a card with the available actions.
          const value = { count: 0 };
          const card = CardFactory.heroCard("Lets talk...", null, [
            {
              type: ActionTypes.MessageBack,
              title: "Say Hello",
              value: value,
              text: "Hello",
            },
          ]);
          await context.sendActivity({ attachments: [card] });
          break;
      }
      await next();
    });
  }

  /**
   * Say hello and @ mention the current user.
   */
  async mentionActivityAsync(context) {
    const TextEncoder = require("html-entities").XmlEntities;

    const mention = {
      mentioned: context.activity.from,
      text: `<at>${new TextEncoder().encode(context.activity.from.name)}</at>`,
      type: "mention",
    };

    const replyActivity = MessageFactory.text(`Hi ${mention.text}`);
    replyActivity.entities = [mention];

    await context.sendActivity(replyActivity);
  }
}

module.exports.BotActivityHandler = BotActivityHandler;

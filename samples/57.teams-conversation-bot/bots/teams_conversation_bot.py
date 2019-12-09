from botframework.connector.auth import MicrosoftAppCredentials
from botbuilder.core import BotFrameworkAdapter, CardFactory, TurnContext, MessageFactory
from botbuilder.core.teams import TeamsActivityHandler, TeamsInfo
from botbuilder.core.teams.teams_activity_extensions import teams_get_channel_id
from botbuilder.schema import CardAction, ConversationParameters, HeroCard, Mention
from botbuilder.schema._connector_client_enums import ActionTypes


class TeamsConversationBot(TeamsActivityHandler):
    def __init__(self, app_id: str, app_password: str):
        self._app_id = app_id
        self._app_password = app_password

    async def on_message_activity(self, turn_context: TurnContext):
        TurnContext.remove_recipient_mention(turn_context.activity)
        turn_context.activity.text = turn_context.activity.text.strip()

        if turn_context.activity.text == "MentionMe":
            await self._mention_activity(turn_context)
            return

        if turn_context.activity.text == "UpdateCardAction":
            await self._update_card_activity(turn_context)
            return

        if turn_context.activity.text == "Delete":
            await self._delete_card_activity(turn_context)
            return

        card = HeroCard(
            title="Welcome Card",
            text="Click the buttons to update this card",
            buttons=[
                CardAction(
                        type=ActionTypes.message_back,
                        title="Update Card",
                        text="UpdateCardAction",
                        value={"count": 0}
                ),
                CardAction(
                        type=ActionTypes.message_back,
                        title="Message all memebers",
                        text="MessageAllMembers"
                )
            ]
        )
        await turn_context.send_activity(MessageFactory.attachment(CardFactory.hero_card(card)))
        return

    async def _mention_activity(self, turn_context: TurnContext):
        mention = Mention(
            mentioned=turn_context.activity.from_property,
            text=f"<at>{turn_context.activity.from_property.name}</at>",
            type="mention"
        )

        reply_activity = MessageFactory.text(f"Hello {mention.text}")
        reply_activity.entities = [Mention().deserialize(mention.serialize())]
        await turn_context.send_activity(reply_activity)

    async def _update_card_activity(self, turn_context: TurnContext):
        data = turn_context.activity.value
        data["count"] += 1

        card = CardFactory.hero_card(HeroCard(
            title="Welcome Card",
            text=f"Updated count - {data['count']}",
            buttons=[
                CardAction(
                    type=ActionTypes.message_back,
                    title='Update Card',
                    value=data,
                    text='UpdateCardAction'
                ),
                CardAction(
                    type=ActionTypes.message_back,
                    title='Message all members',
                    text='MessageAllMembers'

                ),
                CardAction(
                    type= ActionTypes.message_back,
                    title='Delete card',
                    text='Delete'
                )
            ]
            )
        )
        
        updated_activity = MessageFactory.attachment(card)
        updated_activity.id = turn_context.activity.reply_to_id
        await turn_context.update_activity(updated_activity)
    
    async def _message_all_members(self, turn_context: TurnContext):
        teams_channel_id = teams_get_channel_id(turn_context.activity)
        service_url = turn_context.activity.service_url
        creds = MicrosoftAppCredentials(app_id=self._app_id, password=self._app_password)
        conversation_reference = None

        team_members = await TeamsInfo.get_members(turn_context)

        for member in team_members:
            proactive_message = MessageFactory.text(f"Hello {member.given_name} {member.surname}. I'm a Teams conversation bot.")

            ref = TurnContext.get_conversation_reference(turn_context.activity)
            ref.user = member

            async def get_ref(tc1: TurnContext):
                ref2 = TurnContext.get_conversation_reference(t1.activity)
                await BotFrameworkAdapter(t1.adapter).continue_conversation(ref2, lambda t2: await t2.send_activity(proactive_message))

            await BotFrameworkAdapter(turn_context.adapter).create_conversation(ref, )
            
           
    
    async def _delete_card_activity(self, turn_context: TurnContext):
        await turn_context.delete_activity(turn_context.activity.reply_to_id)

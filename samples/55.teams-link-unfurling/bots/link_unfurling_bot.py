from botbuilder.core import CardFactory, TurnContext, MessageFactory
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import Attachment, CardAction, CardImage, HeroCard, ThumbnailCard
from botbuilder.schema.teams import AppBasedLinkQuery, MessagingExtensionAttachment, MessagingExtensionQuery, MessagingExtensionResult, MessagingExtensionResponse
from botbuilder.schema._connector_client_enums import ActionTypes


class LinkUnfurlingBot(TeamsActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        await turn_context.send_activities(MessageFactory.text(f"Echo: {turn_context.activity.text}"))

    async def on_teams_app_based_link_query(self, turn_context: TurnContext, query: AppBasedLinkQuery):
        hero_card = ThumbnailCard(
            title="Thumbnail Card",
            text="query.Url",
            images=[CardImage(url="https://raw.githubusercontent.com/microsoft/botframework-sdk/master/icon.png")]
        )

        attachments = MessagingExtensionAttachment(content_type=CardFactory.content_types.hero_card, content=hero_card)
        result = MessagingExtensionResult(attachment_layout="list", type="result", attachments=[attachments])

        return MessagingExtensionResponse(compose_extension=result)

    async def on_teams_messaging_extension_query( # pylint: disable=unused-argument
        self, turn_context: TurnContext, query: MessagingExtensionQuery):
        if query.command_id == "searchQuery":
            card = HeroCard(
                title="This is a Link Unfurling Sample",
                subtitle="IT will unfurl links from *.botframework.com",
                text="This sample demonstrates how to handle link unfurling in Teams. Plase review the readme for more inforamtion."
            )
            return MessagingExtensionResponse(
                compose_extension=MessagingExtensionResult(
                    attachment_layout="list",
                    type="result",
                    attachments=[
                        MessagingExtensionAttachment(
                            content=card,
                            content_type=CardFactory.content_types.hero_card,
                            preview=Attachment(
                                content=card,
                                content_type=CardFactory.content_types.hero_card
                            )
                        )
                    ]
                )
            )
        
        raise NotImplementedError(f"Invalid CommandId: {query.command_id}")
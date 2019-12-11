# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.


from botbuilder.core import CardFactory, MessageFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import Attachment
from botbuilder.schema.teams import (
    MessagingExtensionAction, 
    MessagingExtensionActionResponse, 
    MessagingExtensionResult,
    TaskModuleContinueResponse,
    TaskModuleTaskInfo
)


class ActionBasedFetchTaskBot(TeamsActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        if turn_context.activity.value:
            answer = turn_context.activity.value["Answer"]
            choices = turn_context.activity.value["Choices"]
            await turn_context.send_activity(MessageFactory.text(f"Your answer '{answer}' and choice {choices}"))
        else:
            await turn_context.send_activities(MessageFactory.text("Hello from the ActionBasedMessagingExtensionFetchTaskBot"))
        
    
    async def on_teams_messaging_extension_fetch_task(self, turn_context: TurnContext, action: MessagingExtensionAction):
        return MessagingExtensionActionResponse(
            task=TaskModuleContinueResponse(
                value=TaskModuleTaskInfo(
                    card=Attachment(
                        content={
                            "body": [
                                {
                                    "text": "Enter text fora question:"
                                },
                                {
                                    "id": "Question",
                                    "placeholder": "Question text here",
                                    "value": "question"
                                }
                            "actions": [ 
                                {
                                    "type": "Action.Submit",
                                    "title": "Submit",
                                    "data": {
                                        "submitLocation": "messagingExtensionFetchTask"
                                    }
                                }
                            ]

                            ]
                        },
                        content_type=CardFactory.content_types.adaptive_card
                    )
                )
            )
        )
    
    async def on_teams_messaging_extension_bot_message_preview_edit_activity(self, turn_context: TurnContext, action: MessagingExtensionAction):
        return MessagingExtensionActionResponse(
            compose_extension=MessagingExtensionResult(
                type="botMessagePreview",
                activity_preview=MessageFactory.attachment(Attachment(
                    content_type=CardFactory.content_types.adaptive_card,
                    content={
                        "body": [
                            {
                                "text": "Adaptive card from task module"
                            },
                            {
                                "text": action.data.question,
                                "id": "Question"
                            },
                            {
                                "id": "Answer",
                                "placeholder": "Answer here...."
                            }
                        ],
                        "actions": [
                            {
                                "type": "Action.Submit",
                                "title": "Submit",
                                "data": {
                                    "submitLocation": "messagingExtensionSubmit"
                                }
                            }
                        ]
                    }
                ))
            )
        )
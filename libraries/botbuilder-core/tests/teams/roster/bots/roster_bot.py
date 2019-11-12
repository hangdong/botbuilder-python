# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from botbuilder.core import MessageFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema.teams import ChannelInfo, TeamInfo, TeamsChannelAccount


class RosterBot(TeamsActivityHandler):
    async def on_message_activty(self, turn_context: TurnContext):
        await turn_context.send_activity(MessageFactory.text(f"Echo: {turn_context.activity.text}"))

        TurnContext.remove_recipient_mention(turn_context.activity)

        if turn_context.text == "show members":
            await self._show_members_activity(turn_context)
            return
        elif turn_context.text == "show channels":
            await self._show_channels_activity(turn_context)
            return
        elif turn_context.text == "show details":
            await self._show_details_activity(turn_context)
            return
        else:
            await turn_context.send_activity
            (
                MessageFactory.text
                                (
                                    """Invalid command. Type \"Show channels\" to see a channel list.
                                    Type \"Show members\" to see a list of members in a team.
                                    "Type \"show group chat members\" to see members in a group chat."""
                                )
            )
            return

    async def _show_members_activity(turn_context: TurnContext):
        team_id = TurnContext.teams_get_team_info().id

        
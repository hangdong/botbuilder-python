# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from botbuilder.core import MessageFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler, TeamsInfo
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
        teams_channel_account = await TeamsInfo.get_members(turn_context)
        await turn_context.send_activity(MessageFactory.text(f"There are {len(teams_channel_account)} members are currently in the team"))

        messages = [f"{member.aad_group_id} --> {member.name} --> {member.user_principal_name}" for member in teams_channel_account]
        a
    
    async def _show_channels_activity(turn_context: TurnContext):
        pass
    
    async def _show_details_activity(turn_context: TurnContext):
        team_id = TurnContext.teams_get_team_info().id
        team_details = await TeamsInfo.get_team_details(turn_context, team_id)
        reply_activity = MessageFactory.text(f"""The team name is {team_details.name}. 
                                                The team ID is {team_details.id}. 
                                                The AADGroupID is {team_details.aad_group_id}""")
        await turn_context.send_activity(reply_activity)

    async def _send_in_batches(turn_context: TurnContext, messages: []):
        batch = []
        for message in messages:
            batch.append(message)
            if len(batch) == 10:
                await turn_context.send_activity(MessageFactory.text(batch.join("<br>")))
                batch = []
        if len(batch) > 0:
            await turn_context.send_activity(MessageFactory.text(batch.join("<br>")))
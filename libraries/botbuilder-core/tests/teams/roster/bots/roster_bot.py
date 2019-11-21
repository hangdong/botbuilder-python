# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from botbuilder.core import MessageFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema.teams import TeamsChannelAccount
from typing import List


class RosterBot(TeamsActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext):
        await turn_context.send_activity(MessageFactory.text(f"Echo: {turn_context.activity.text}"))

        TurnContext.remove_mention_text(turn_context.activity)
        if turn_context.activity.text == "show members":
            await self._show_members_activity(turn_context)
        elif turn_context.activity.text == "show channels":
            await self._show_channels_activity(turn_context)
        elif turn_context.activity.text == "show details":
            await self._show_details_activity(turn_context)
        else:
            await turn_context.send_activity(
                                            MessageFactory.text(
                                                """Invalid command. Type \"Show channels\" to see a channel list. 
                                                Type \"Show members\" to see a list of members in a team. " +
                                                "Type \"show group chat members\" to see members in a group chat."""
                                                )
                                            )
        return
    
    async def _show_members_activity(self, turn_context: TurnContext)
        team_id = turn_context.activity.teams_get_team_info().id

        team_details = await TeamsInfo.get_team_details_activity(turn_context, team_id)

        reply_activity = MessageFactory.text(f"""The team name is {team_deails.name}. 
                                                The team ID is {team_details.id}. 
                                                The AADGroupId is {team_details.aad_group_id}""")
        await turn_context.send_activity(reply_activity)
    
    async def _show_channels_activity(self, turn_context: TurnContext):
        team_id = turn_context.activity.TeamsGetTeamInfo().id

        channels = await TeamsInfo.GetTeamsChannelsActivity(turn_context, team_id)

        reply_activity = MessageFactory.text(f"Total of {len(channels)} channels are currently in team")

        messages = [ f"{channel.id} -> {channel.name}" for channel in channels]

        await self._send_in_batches_activity(turn_context, messages)
        
    async def _show_members_activity(self, turn_context: TurnContext, teams_channel_accounts: List[TeamsChannelAccount]):
        reply_activity = MessageFactory.text(f"Total of {len(teams_channel_accounts)} members are currently in team")

        await turn_context.send_activity(reply_activity)

        messages = [f"{member.name} -> {member.aad_group_id} -> {member.user_principal_name}" for member in teams_channel_accounts]

        await self._send_in_batches_activity(turn_context, messages)
    
    async def _send_in_batches_activity(turn_context: TurnContext, messages: List[str]):
        batch = []
        for msg in messages:
            batch.append(msg)
            if len(batch) == 10:
                await turn_context.send_activity(MessageFactory.text("<br>".join(batch)))
                batch = []
        
        if len(batch) > 0:
            await turn_context.send_activity(MessageFactory.text("<br>".join(batch)))
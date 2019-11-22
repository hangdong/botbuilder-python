# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from botbuilder.core import TurnContext
from botbuilder.schema.teams import ChannelInfo, TeamsChannelData, TeamDetails
from botbuilder.connector import BotBuilder

from typing import List

class TeamsInfo:
    @staticmethod
    async def get_team_details(self, turn_context: TurnContext, team_id: str = None) -> TeamDetails:
        team_info = turn_context.activity.teams_get_team_info()
        t = team_id or (team_id.id if team_info and team_info.id else None)
        if not t:
            raise Exception("This method is only valid within the scope of MS Teams team")
        return await self._get_teams_connector_client(turn_context).teams.fetch_team_details_activity(t)
    
    @staticmethod
    async def get_team_channel_activity(self, turn_context: TurnContext, team_id: str = None) -> List[ChannelInfo]:
        team_info = turn_context.activity.teams_get_team_info()
        t = team_id or (team_id.id if team_info and team_info.id else None)
        if not t:
            raise Exception("This method is only valid within the scope of MS Teams team")
        channel_list = await self._get_teams_connector_client(turn_context).teams.fetch_channel_list_activity(t)
        return channel_list.conversations

    @staticmethod
    async def get_team_members_activity(self, turn_context: TurnContext, team_id: str = None) -> List[TeamsChannelAccount]:
        team_info = turn_context.activity.teams_get_team_info()
        t = team_id or (team_id.id if team_info and team_info.id else None)
        if not t:
            raise Exception("This method is only valid within the scope of MS Teams team")
        return self._get_members_activity(self._get_connector_client(turn_context), t)
    
    @staticmethod
    async def get_members_activity(self, turn_context: TurnContext):
        team_info = turn_context.activity.teams_get_team_info()
        if team_info and team_info.id:
            return self._get_team_members_activity(turn_context, team_info.id)
        else:
            conversation_id = turn_context.activity.conversation.id if turn_context.activity and turn_context.activity.conversation else None
            return self._get_members_activity(self._get_connector_client(turn_context), conversation_id)
    
    @staticmethod
    async def get_members_activity(self, connector_client: ConnectorClient, conversation_id: str) -> List[TeamsChannelAccount]:
        if not conversation_id:
            raise Exception("The GetMembers operation needs a valid conversation ID.")
        
        team_members = await connector_client.conversations.get_conversation_members_activity(conversation_id)
        teams_channel_accounts = [TeamsChannelAccount(**members) for members in team_members]
        return teams_channel_accounts
    
    @staticmethod
    def get_team_id(self, turn_context: TurnContext):
        if not turn_context:
            raise Exception("Missing context parameter")
        
        if not turn_context.activity:
            raise Exception("Missing activyt on turn_context")
        
        channel_data = turn_context.activity.channel_data
        team = channel_data.team if channel_data and channel_data.team else None
        team_id = team.id if team and type(team.id) == 'str' else None
        return team_id
    
    @staticmethod
    def get_connector_client(self, turn_context: TurnContext) -> ConnectorClient:
        if not turn_context.adapter or (not 'createConnectorClient' in turn_context.adapter):
            raise Exception("This method requries a connector client")
        return turn_context.adapter.create_connector_client(turn_context.activity.service_url)
    
    @staticmethod
    def get_teams_connector_client(self, turn_context: TurnContext) -> TeamsConnectorClient:
        connector_client = self._get_connector_client(turn_context)
        return TeamsConnectorClient(connector_client.credentials, {"base_url": turn_context.activity.service_url})
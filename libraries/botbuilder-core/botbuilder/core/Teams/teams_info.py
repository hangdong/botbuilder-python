from botbuilder.core import TurnContext
from botbuilder.schema import ChannelAccount
from botbuilder.schema.teams import ChannelInfo, TeamsChannelAccount, TeamInfo, TeamDetails
from botframework.connector import ConnectorClient
from botframework.connector.teams import TeamsConnectorClient
from typing import List

class TeamsInfo:
    @staticmethod
    async def get_team_details(turn_context: TurnContext, team_id: str = None) -> TeamDetails:
        t = team_id if team_id else self._get_team_id(turn_context)
        if not t:
            raise TypeError("This method is only valid within the scope of a MS Teams team.")
        
        return await self._get_teams_connector_client(turn_context).teams.fetch_team_details(t)
    
    @staticmethod
    async def get_team_channels(turn_context: TurnContext, team_id: str = None) -> List[ChannelInfo]:
        t = team_id if team_id else self._get_team_id(turn_context)
        if not t:
            raise TypeError("This method is only valid within the scope of a MS Teams team.")
        
        channel_list = await self._get_teams_connector_client(turn_context).teams.fetch_channel_list(t)
        return channel_list.conversations
    
    @staticmethod
    async def get_members(turn_context: TurnContext) -> List[TeamsChannelAccount]:
        team_id = self._get_team_id(turn_context)
        if team_id:
            return await self.get_team_members(turn_context, team_id)
        else:
            conversation = turn_context.activity.conversation
            conversation_id = conversation.id if conversation and conversation.id else None
            return await self._get_members_internal(self._get_connector_client(turn_context), conversation_id)
    
    @staticmethod
    async def get_team_members(turn_Context: TurnContext, team_id: str = None) -> List[TeamsChannelAccount]:
        t = team_id if team_id else self._get_team_id(turn_context)
        if not t:
            raise TypeError("This method is only valid within the scope of a MS Teams team.")
        return self._get_members_internal(self._get_connector_client(turn_context), t)
    
    @staticmethod
    async def _get_members_internal(connector_client: ConnectorClient, conversation_id: str) -> List[TeamsChannelAccount]:
        t = team_id if team_id else self._get_team_id(turn_context)
        team_members = await connector_client.conversations.get_conversation_members(conversation_id)
        teams_channel_account_list = List[TeamsChannelAccount]
        for member in team_members:
            teams_channel_account_list.append(TeamsChannelAccount(**member))
        return teams_channel_account_list
    
    @staticmethod
    async def _get_team_id(turn_context: TurnContext) -> str:
        if not turn_context:
            raise TypeError("Missing turn_context parameter")

        if not turn_context.activity:
            raise TypeError("Missing activity on turn_context")

        channel_data = TeamsChannelData(**turn_context.activity.channel_data)
        team = channel_data.team if channel_data and channel_data.team else None
        team_id = team.id if team and type(team.id) == 'str' else None
        return team_id

    @staticmethod
    async def _get_connector_client(turn_context: TurnContext) -> ConnectorClient:
        if not context.adapter or "createConnectorClient" in context.adapter:
            raise TypeError("This method requries a connector client")
        
        return BotFrameworkAdapter(**context.adapter).createConnectorClient(turn_context.activity.service_url)
    
    @staticmethod
    async def _get_teams_connector_client(turn_context: TurnContext) -> TeamsConnectorClient:
        pass
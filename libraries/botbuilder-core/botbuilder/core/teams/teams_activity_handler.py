# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from http import HTTPStatus
from botbuilder.schema import ActivityTypes, ChannelAccount
from botbuilder.core.turn_context import TurnContext
from botbuilder.core import ActivityHandler, MessageFactory, InvokeResponse
from botbuilder.schema.teams import (
    TeamInfo,
    ChannelInfo,
    TeamsChannelData,
    TeamsChannelAccount,
)
from botframework.connector import Channels

class TeamsActivityHandler(ActivityHandler):
    async def on_turn(self, turn_context: TurnContext):
        if turn_context is None:
            raise TypeError("ActivityHandler.on_turn(): turn_context cannot be None.")

        if hasattr(turn_context, "activity") and turn_context.activity is None:
            raise TypeError(
                "ActivityHandler.on_turn(): turn_context must have a non-None activity."
            )

        if (
            hasattr(turn_context.activity, "type")
            and turn_context.activity.type is None
        ):
            raise TypeError(
                "ActivityHandler.on_turn(): turn_context activity must have a non-None type."
            )

        if turn_context.activity.type == ActivityTypes.invoke:
            invoke_response = await self.on_invoke_activity(turn_context)
            if invoke_response and not turn_context.turn_state[BotFrameworkAdapter()._INVOKE_RESPONSE_KEY]:
                await turn_context.send_activity(Activity(value=invoke_response, type=ActivityTypes.InvokeResponse))
            return
        else:
            await super().on_turn(turn_context)
            return

    async def on_invoke_activity(self, turn_context: TurnContext):
        try:
            if not turn_context.activity.name and turn_context.activity.channel_id == Channels.Msteams:
                return await self.on_teams_card_action_invoke_activity(turn_context)
            else:
                if turn_context.activity.name == "signin/verifyState":
                    await self.on_teams_signin_verify_state_activity(turn_context)
                    return self._create_invoke_response()
                elif turn_context.activity.name == "fileConsent/invoke":
                    return await self.on_teams_file_consent_activity(turn_context)
                elif turn_context.activity.name == "actionableMessage/executeAction":
                    await on_teams_o365_connector_card_action_activity(turn_context)
                    return self._create_invoke_response()
                elif turn_context.activity.name == "composeExtension/queryLink":
                    return self._create_invoke_response(await self.on_teams_app_based_link_query_activity(turn_context, turn_context.activity.value))
                elif turn_context.activity.name == "composeExtension/query":
                    return self._create_invoke_response(await self.on_teams_messaging_extension_query_activity(turn_context, turn_context.activity.value))
                elif turn_context.activity.name == "composeExtension/selectItem":
                    return self._create_invoke_response(await self.on_teams_messaging_extension_select_item_activity(turn_context, turn_context.activity.value))
                elif turn_context.activity.name == "composeExtension/submitAction":
                    return self._create_invoke_response(await self.on_teams_messaging_etension_submit_action_dispatch_activity(turn_context, turn_context
                    .activity.value))
                elif turn_context.activity.name == "composeExtension/fetchTask":
                    return self._create_invoke_response(await self.on_teams_messaging_extension_fetch_task_activity(turn_context, turn_context
                    .activity.value))
                elif turn_context.activity.name == "composeExtension/querySettingUrl":
                    return self._create_invoke_response(await on_teams_messaging_extension_configuration_query_settings_url_activity(turn_context, turn_context.activity.value))
                elif turn_context.activity.name == "composeExtension/setting":
                    await on_Teams_messaging_extension_configuration_setting_activity(turn_context, turn_context.turn_context.activity.value)
                    return self._create_invoke_response()
                elif turn_context.activity.name == "composeExtension/onCardButtonClicked":
                    await self.on_teams_messaging_extension_card_button_clicked_activity(turn_context, turn_context.activity.value)
                    return self._create_invoke_response()
                elif turn_context.activity.name == "task/fetch":
                    return self._create_invoke_response(await self.on_Teams_task_module_fetch_activity(turn_context, turn_context.activity.value))
                elif turn_context.activity.name == "task/submit":
                    return self._create_invoke_response(await self.on_teams_task_module_submit_activity(turn_context, turn_context.activity.value))
                else:
                    raise InvokeResponseException(HTTPStatus.NotImplemented)
        except(Exception e):
            raise self._create_invoke_response(e)

    async def on_teams_file_consent_activity(self, turn_context: TurnContext, file_consent_card_response):
        if file_consent_card_response.action == "accept":
            await self.on_teams_file_consent_accept_activity(turn_context, file_consent_card_response)
            return self._create_invoke_response()
        elif file_consent_card_response.action == "decline":
            await on_teams_file_consent_decline_activity(turn_context, file_consent_card_response)
            return self._create_invoke_response()
        else:
            raise InvokeResponseException(HTTPStatus.BadRequest, f"{file_consent_card_response.action} is not a supported Action.")

    async def on_teams_messaging_extension_submit_action_dispatch_activity(self, turn_context: TurnContext, action):
        if action:
            if action.bot_message_preview_action == "edit":
                return await self.on_teams_messaging_extension_bot_message_preview_edit_activity(turn_context, action)
            elif action.bot_message_preview_action == "send":
                return await self.on_teams_messaging_extension_bot_message_send_activity(turn_context, action)
            else:
                raise InvokeResponseException(HTTPStatus.BadRequest, f"{action.bot_message_preview_action} is not a supported BotMessagePreviewAction")
        else:
            return await on_teams_messaging_extension_submit_action_activity(turn_context, action)

    async def on_conversation_update_activity(self, turn_context: TurnContext):
        if turn_context.activity.channel_id == Channels.ms_teams:
            channel_data = TeamsChannelData(**turn_context.activity.channel_data)

            if turn_context.activity.members_added:
                return await self.on_teams_members_added_dispatch_activity(
                    turn_context.activity.members_added, channel_data.team, turn_context
                )

            if turn_context.activity.members_removed:
                return await self.on_teams_members_removed_dispatch_activity(
                    turn_context.activity.members_removed,
                    channel_data.team,
                    turn_context,
                )

            if channel_data:
                if channel_data.event_type == "channelCreated":
                    return await self.on_teams_channel_created_activity(
                        channel_data.channel, channel_data.team, turn_context
                    )
                if channel_data.event_type == "channelDeleted":
                    return await self.on_teams_channel_deleted_activity(
                        channel_data.channel, channel_data.team, turn_context
                    )
                if channel_data.event_type == "channelRenamed":
                    return await self.on_teams_channel_renamed_activity(
                        channel_data.channel, channel_data.team, turn_context
                    )
                if channel_data.event_type == "teamRenamed":
                    return await self.on_teams_team_renamed_activity(
                        channel_data.team, turn_context
                    )
                return await super().on_conversation_update_activity(turn_context)

        return await super().on_conversation_update_activity(turn_context)

    async def on_teams_channel_created_activity(  # pylint: disable=unused-argument
        self, channel_info: ChannelInfo, team_info: TeamInfo, turn_context: TurnContext
    ):
        return

    async def on_teams_team_renamed_activity(  # pylint: disable=unused-argument
        self, team_info: TeamInfo, turn_context: TurnContext
    ):
        return

    async def on_teams_members_added_dispatch_activity(  # pylint: disable=unused-argument
        self,
        members_added: [ChannelAccount],
        team_info: TeamInfo,
        turn_context: TurnContext,
    ):
        """
        team_members = {}
        team_members_added = []
        for member in members_added:
            if member.additional_properties != {}:
                team_members_added.append(TeamsChannelAccount(member))
            else:
                if team_members == {}:
                    result = await TeamsInfo.get_members_async(turn_context)
                    team_members = { i.id : i for i in result }

                if member.id in team_members:
                    team_members_added.append(member)
                else:
                    newTeamsChannelAccount = TeamsChannelAccount(
                        id=member.id,
                        name = member.name,
                        aad_object_id = member.aad_object_id,
                        role = member.role
                        )
                    team_members_added.append(newTeamsChannelAccount)

        return await self.on_teams_members_added_activity(teams_members_added, team_info, turn_context)
        """
        for member in members_added:
            new_account_json = member.__dict__
            del new_account_json["additional_properties"]
            member = TeamsChannelAccount(**new_account_json)
        return await self.on_teams_members_added_activity(members_added, turn_context)

    async def on_teams_members_added_activity(
        self, teams_members_added: [TeamsChannelAccount], turn_context: TurnContext
    ):
        for member in teams_members_added:
            member = ChannelAccount(member)
        return super().on_members_added_activity(teams_members_added, turn_context)

    async def on_teams_members_removed_dispatch_activity(  # pylint: disable=unused-argument
        self,
        members_removed: [ChannelAccount],
        team_info: TeamInfo,
        turn_context: TurnContext,
    ):
        teams_members_removed = []
        for member in members_removed:
            new_account_json = member.__dict__
            del new_account_json["additional_properties"]
            teams_members_removed.append(TeamsChannelAccount(**new_account_json))

        return await self.on_teams_members_removed_activity(
            teams_members_removed, turn_context
        )

    async def on_teams_members_removed_activity(
        self, teams_members_removed: [TeamsChannelAccount], turn_context: TurnContext
    ):
        members_removed = [ChannelAccount(i) for i in teams_members_removed]
        return super().on_members_removed_activity(members_removed, turn_context)

    async def on_teams_channel_deleted_activity(  # pylint: disable=unused-argument
        self, channel_info: ChannelInfo, team_info: TeamInfo, turn_context: TurnContext
    ):
        return  # Task.CompleteTask

    async def on_teams_channel_renamed_activity(  # pylint: disable=unused-argument
        self, channel_info: ChannelInfo, team_info: TeamInfo, turn_context: TurnContext
    ):
        return  # Task.CompleteTask

    async def on_teams_team_reanamed_async(  # pylint: disable=unused-argument
        self, team_info: TeamInfo, turn_context: TurnContext
    ):
        return  # Task.CompleteTask

    @staticmethod
    def _create_invoke_response(body: object = None) -> InvokeResponse:
        return InvokeResponse(status=int(HTTPStatus.OK), body=body)

    class _InvokeResponseException(Exception):
        def __init__(self, status_code: HTTPStatus, body: object = None):
            super().__init__()
            self._status_code = status_code
            self._body = body

        def create_invoke_response(self) -> InvokeResponse:
            return InvokeResponse(status=int(self._status_code), body=self._body)
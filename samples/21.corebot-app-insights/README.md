# CoreBot with Application Insights

Bot Framework v4 core bot sample.

This bot has been created using [Bot Framework](https://dev.botframework.com), it shows how to:

- Use [LUIS](https://www.luis.ai) to implement core AI capabilities
- Implement a multi-turn conversation using Dialogs
- Handle user interruptions for such things as `Help` or `Cancel`
- Prompt for and validate requests for information from the user
- Use [Application Insights](https://docs.microsoft.com/azure/azure-monitor/app/cloudservices) to monitor your bot

## Prerequisites

### Install Python 3.6

### Overview

This bot uses [LUIS](https://www.luis.ai), an AI based cognitive service, to implement language understanding 
and [Application Insights](https://docs.microsoft.com/azure/azure-monitor/app/cloudservices), an extensible Application Performance Management (APM) service for web developers on multiple platforms.

### Create a LUIS Application to enable language understanding

LUIS language model setup, training, and application configuration steps can be found [here](https://docs.microsoft.com/azure/bot-service/bot-builder-howto-v4-luis?view=azure-bot-service-4.0&tabs=cs).

If you wish to create a LUIS application via the CLI, these steps can be found in the [README-LUIS.md](README-LUIS.md).

### Add Application Insights service to enable the bot monitoring
Application Insights resource creation steps can be found [here](https://docs.microsoft.com/azure/azure-monitor/app/create-new-resource).

## Running the sample
- Run `pip install -r requirements.txt` to install all dependencies
- Update LuisAppId, LuisAPIKey and LuisAPIHostName in `config.py` with the information retrieved from the [LUIS portal](https://www.luis.ai)
- Update AppInsightsInstrumentationKey in `config.py`
- Run `python app.py`
- Alternatively to the last command, you can set the file in an environment variable with `set FLASK_APP=app.py` in windows (`export FLASK_APP=app.py` in mac/linux) and then run `flask run --host=127.0.0.1 --port=3978`


## Testing the bot using Bot Framework Emulator

[Bot Framework Emulator](https://github.com/microsoft/botframework-emulator) is a desktop application that allows bot developers to test and debug their bots on localhost or running remotely through a tunnel.

- Install the Bot Framework Emulator version 4.3.0 or greater from [here](https://github.com/Microsoft/BotFramework-Emulator/releases)

### Connect to the bot using Bot Framework Emulator

- Launch Bot Framework Emulator
- File -> Open Bot
- Enter a Bot URL of `http://localhost:3978/api/messages`


## Further reading

- [Bot Framework Documentation](https://docs.botframework.com)
- [Bot Basics](https://docs.microsoft.com/azure/bot-service/bot-builder-basics?view=azure-bot-service-4.0)
- [Dialogs](https://docs.microsoft.com/azure/bot-service/bot-builder-concept-dialog?view=azure-bot-service-4.0)
- [Gathering Input Using Prompts](https://docs.microsoft.com/azure/bot-service/bot-builder-prompts?view=azure-bot-service-4.0&tabs=csharp)
- [Activity processing](https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-concept-activity-processing?view=azure-bot-service-4.0)
- [Azure Bot Service Introduction](https://docs.microsoft.com/azure/bot-service/bot-service-overview-introduction?view=azure-bot-service-4.0)
- [Azure Bot Service Documentation](https://docs.microsoft.com/azure/bot-service/?view=azure-bot-service-4.0)
- [Language Understanding using LUIS](https://docs.microsoft.com/azure/cognitive-services/luis/)
- [Channels and Bot Connector Service](https://docs.microsoft.com/azure/bot-service/bot-concepts?view=azure-bot-service-4.0)
- [Application insights Overview](https://docs.microsoft.com/azure/azure-monitor/app/app-insights-overview)
- [Getting Started with Application Insights](https://github.com/Microsoft/ApplicationInsights-aspnetcore/wiki/Getting-Started-with-Application-Insights-for-ASP.NET-Core)
- [Filtering and preprocessing telemetry in the Application Insights SDK](https://docs.microsoft.com/azure/azure-monitor/app/api-filtering-sampling)

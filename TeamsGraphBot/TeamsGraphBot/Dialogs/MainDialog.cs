using Microsoft.Bot.Builder;
using Microsoft.Bot.Builder.Dialogs;
using Microsoft.Bot.Builder.Teams;
using Microsoft.Bot.Schema;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace TeamsGraphBot.Dialogs
{
	public class MainDialog : ComponentDialog
	{
		protected readonly ILogger _logger;

		public MainDialog(ILogger<MainDialog> logger, IConfiguration configuration) : base(nameof(MainDialog))
		{
			_logger = logger;

			AddDialog(new OAuthPrompt(
				nameof(OAuthPrompt),
				new OAuthPromptSettings
				{
					ConnectionName = configuration["ConnectionName"],
					Text = "Please login",
					Title = "Login",
					Timeout = 300000
				}));



			AddDialog(new WaterfallDialog(nameof(WaterfallDialog), new WaterfallStep[]
			{
				PromptStepAsync,
				DisplayContextInfoStepAsync
			}));

			InitialDialogId = nameof(WaterfallDialog);

		}

		private async Task<DialogTurnResult> PromptStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			return await stepContext.BeginDialogAsync(nameof(OAuthPrompt), null, cancellationToken);
		}



		private async Task<DialogTurnResult> DisplayContextInfoStepAsync(WaterfallStepContext stepContext, CancellationToken cancellationToken)
		{
			if (stepContext.Result != null)
			{
				var tokenResponse = stepContext.Result as TokenResponse;
				if (tokenResponse?.Token != null)
				{
					var teamsContext = stepContext.Context.TurnState.Get<ITeamsContext>();

					if (teamsContext != null) // the bot is used inside MS Teams
					{
						if (teamsContext.Team != null) // inside team
						{
							var team = teamsContext.Team;
							var teamDetails = await teamsContext.Operations.FetchTeamDetailsWithHttpMessagesAsync(team.Id);
							var token = tokenResponse.Token;

							var graphClient = new GraphServiceClient(
								new DelegateAuthenticationProvider(
								requestMessage =>
								{
									// Append the access token to the request.
									requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);

									// Get event times in the current time zone.
									requestMessage.Headers.Add("Prefer", "outlook.timezone=\"" + TimeZoneInfo.Local.Id + "\"");

									return Task.CompletedTask;
								})
							);

							var siteInfo = await graphClient.Groups[teamDetails.Body.AadGroupId].Sites["root"].Request().GetAsync();

							await stepContext.Context.SendActivityAsync(MessageFactory.Text($"Site Id: {siteInfo.Id}, Site Title: {siteInfo.DisplayName}, Site Url: {siteInfo.WebUrl}"), cancellationToken).ConfigureAwait(false);
						}
						else // private or group chat
						{
							await stepContext.Context.SendActivityAsync(MessageFactory.Text($"We're in MS Teams but not in Team"), cancellationToken).ConfigureAwait(false);
						}
					}
					else // outside MS Teams
					{
						await stepContext.Context.SendActivityAsync(MessageFactory.Text("We're not in MS Teams context"), cancellationToken).ConfigureAwait(false);
					}
				}
			}

			return await stepContext.EndDialogAsync();
		}
	}
}

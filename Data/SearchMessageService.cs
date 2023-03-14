namespace SearchMessage.Data
{
    using System;
    using System.Text.Json;
    using System.Text.Json.Serialization;

    using Microsoft.Graph;

    public class SearchMessageService
    {
        private readonly GraphServiceClient graphClient;

        public SearchMessageService(GraphServiceClient graphClient)
        {
            this.graphClient = graphClient;
        }

        public async Task<IEnumerable<SearchMessageDto>> SearchMessagesAsync(string value)
        {
            var requests = new List<Task<ISearchEntityQueryCollectionPage>>
            {
                CreateSearchQueryObject(EntityType.Message, value),
                CreateSearchQueryObject(EntityType.ChatMessage, value),
            };
            var results = await Task.WhenAll(requests);
            var searchHits = results.SelectMany(res => res.Where(res => res.HitsContainers != null)
                                        .SelectMany(res => res.HitsContainers.Where(res => res.Hits != null)
                                            .SelectMany(res => res.Hits)));
            return searchHits.Select(sr => new SearchMessageDto
            {
                Channel = GetChannel(sr.Resource),
                Rank = sr.Rank,
                ReceivedDate = GetReceivedDate(sr.Resource),
                ReceivedFrom = GetReceivedFrom(sr.Resource),
                ReceivedEmail = GetReceivedEmail(sr.Resource),
                Subject = GetSubject(sr.Resource),
                Summary = sr.Summary
            });

            Task<ISearchEntityQueryCollectionPage> CreateSearchQueryObject(EntityType entityType, string value)
            {
                return graphClient.Search.Query(new List<SearchRequestObject>
                {
                    new SearchRequestObject
                    {
                        EntityTypes = new List<EntityType>
                        {
                            entityType
                        },
                        Query = new SearchQuery
                        {
                            QueryString = value
                        }
                    }
                }).Request().PostAsync();
            }
        }

        private static string GetChannel(Entity resource)
        {
            if (resource == null)
            {
                return string.Empty;
            }

            return resource is Message ? "Mail" : resource is ChatMessage ? "Teams chat" : string.Empty;
        }

        private static DateTime? GetReceivedDate(Entity resource)
        {
            if (resource == null)
            {
                return null;
            }

            return resource is Message message ? (message.ReceivedDateTime?.UtcDateTime) :
                resource is ChatMessage chatMessage ? (chatMessage.CreatedDateTime?.UtcDateTime) : null;
        }

        private static string? GetReceivedFrom(Entity resource)
        {
            if (resource == null)
            {
                return null;
            }

            return resource is Message message ? GetReceivedDetailFromMessage(message).name :
                resource is ChatMessage chatMessage ? GetReceivedDetailFromChat(chatMessage).name : null;
        }

        private static string? GetReceivedEmail(Entity resource)
        {
            if (resource == null)
            {
                return null;
            }

            return resource is Message message ? GetReceivedDetailFromMessage(message).emailAddress :
                resource is ChatMessage chatMessage ? GetReceivedDetailFromChat(chatMessage).emailAddress : null;
        }

        private static (string name, string emailAddress) GetReceivedDetailFromChat(ChatMessage chatMessage)
        {
            var chatDetails = chatMessage.From.AdditionalData["emailAddress"];
            if (chatDetails == null)
            {
                return (string.Empty, string.Empty);
            }

            var chatFromDetail = DeserializeAnonymousType(chatDetails, new { name = "", address = "" });
            if (chatFromDetail == null)
            {
                return (string.Empty, string.Empty);
            }

            return (chatFromDetail.name, chatFromDetail.address);
        }

        private static (string name, string emailAddress) GetReceivedDetailFromMessage(Message message)
        {
            return (message.From?.EmailAddress?.Name ?? string.Empty, message.From?.EmailAddress?.Address ?? string.Empty);
        }

        private static T? DeserializeAnonymousType<T>(object jsonObject, T anonymousType)
            where T : class
        {
            if (anonymousType == null)
            {
                return null;
            }

            string jsonString = jsonObject.ToString() ?? string.Empty;
            return JsonSerializer.Deserialize<T>(jsonString);
        }

        private static string? GetSubject(Entity resource)
        {
            if (resource == null)
            {
                return null;
            }

            return resource is Message message ? (message.Subject) : null;
        }
    }
}

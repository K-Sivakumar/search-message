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
                CreateSearchQueryObject(EntityType.Event, value),
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

            return resource is Message ? "Message" : resource is ChatMessage ? "Teams chat" : string.Empty;
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

            return resource is Message message ? (message.From?.EmailAddress?.Address) :
                resource is ChatMessage chatMessage && 
                    chatMessage.From.AdditionalData["emailAddress"] is object chatDetails &&
                    chatDetails != null ? DeserializeAnonymousType(chatDetails.ToString(), new { name = "", address = "" }).address : null;
        }

        private static T DeserializeAnonymousType<T>(string json, T anonymousType) => JsonSerializer.Deserialize<T>(json);

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

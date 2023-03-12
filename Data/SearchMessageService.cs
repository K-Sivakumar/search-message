namespace SearchMessage.Data
{
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
                Rank = sr.Rank,
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
    }
}

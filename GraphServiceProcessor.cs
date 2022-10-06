using Microsoft.Graph;

namespace Graph
{
    public class GraphServiceProcessor
    {
        private GraphServiceClient Client;
        public GraphServiceProcessor(GraphServiceClient client)
        {
            Client = client;
        }
        public async Task<IListItemsCollectionPage> GetListItems()
        {
            var listItems = await GetSite().Lists["SkippedCustomers-Dev"].Items.Request().GetAsync();
            return listItems;
        }

        public async Task AddTestListItem()
        {
            var item = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>(){
                        {"Title", "Hello from Rolands!"}
                    }
                }
            };
            await GetSite().Lists["Test"].Items.Request().AddAsync(item);
        }

        public async Task<ISiteDrivesCollectionPage> GetDrives(){
            return await GetSite().Drives.Request().GetAsync();
        }

        private ISiteRequestBuilder GetSite()
        {
            return Client.Sites["atea.sharepoint.com,643b8072-6a5b-4c9c-b633-ab21e514a62e,82f6f84e-2b9e-4f7f-ac5b-9b5218a4b301"];
        }
    }
}
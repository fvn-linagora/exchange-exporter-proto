namespace EchangeExporterProto
{
    using System;
    using System.Linq;
    using System.Collections.Generic;

    using Microsoft.Exchange.WebServices.Data;

    static class PagedItemsSearch
    {
        internal static IEnumerable<T> PageSearchItems<T>(ExchangeService service, FolderId folderId, int pageSize, PropertySet properties, PropertyDefinition sortBy) where T: Item
        {
            // int pageSize = 5;
            int offset = 0;

            // Request one more item than your actual pageSize.
            // This will be used to detect a change to the result
            // set while paging.
            ItemView view = new ItemView(pageSize + 1, offset);

            // view.PropertySet = new PropertySet(ItemSchema.Subject);
            view.PropertySet = properties;
            // view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            view.OrderBy.Add(sortBy, SortDirection.Descending);
            view.Traversal = ItemTraversal.Shallow;

            IEnumerable<T> res = new List<T>();

            bool moreItems = true;
            ItemId anchorId = null;
            while (moreItems)
            {
                try
                {
                    FindItemsResults<Item> results = service.FindItems(folderId, view);
                    moreItems = results.MoreAvailable;

                    if (moreItems && anchorId != null)
                    {
                        // Check the first result to make sure it matches
                        // the last result (anchor) from the previous page.
                        // If it doesn't, that means that something was added
                        // or deleted since you started the search.
                        if (results.Items.FirstOrDefault<Item>()?.Id?.ToString() != anchorId.ToString())
                        {
                            Console.WriteLine("The collection has changed while paging. Some results may be missed.");
                        }
                    }

                    if (moreItems)
                        view.Offset += pageSize;

                    // anchorId = results.Items.Last<Item>().Id;
                    res = results.Items.Cast<T>();
                    anchorId = res.LastOrDefault()?.Id;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception while paging results: {0}", ex.Message);
                }
                // Because you’re including an additional item on the end of your results
                // as an anchor, you don't want to display it.
                // Set the number to loop as the smaller value between
                // the number of items in the collection and the page size.
                // int displayCount = results.Items.Count > pageSize ? pageSize : results.Items.Count;

                foreach (var item in res)
                    yield return item;
            }
        }
    }
}

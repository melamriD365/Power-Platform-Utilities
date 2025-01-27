using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;

namespace PP_UTILITIES
{
    public class Plg_FuzzySearch : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            try
            {
                if (!context.InputParameters.Contains("meaf_entity") || !(context.InputParameters["meaf_entity"] is string entityName))
                    throw new InvalidPluginExecutionException("Required input parameter 'meaf_entity' is missing or invalid (string).");
                if (!context.InputParameters.Contains("meaf_searchTerm") || !(context.InputParameters["meaf_searchTerm"] is string searchTerm))
                    throw new InvalidPluginExecutionException("Required input parameter 'meaf_searchTerm' is missing or invalid (string).");
                if (!context.InputParameters.Contains("meaf_selectColumns") || !(context.InputParameters["meaf_selectColumns"] is string[] selectColumns))
                    throw new InvalidPluginExecutionException("Required input parameter 'meaf_selectColumns' is missing or invalid (string[]).");

                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var orgService = serviceFactory.CreateOrganizationService(context.UserId);

                var optionsDict = new Dictionary<string, string>
                {
                    { "fuzzymatchingenabled", "true" },
                    { "grouprankingenabled", "true" }
                };

                var entitiesArray = new[]
                {
                    new
                    {
                        name = entityName,
                        selectcolumns = new List<string>(selectColumns)
                    }
                };

                var searchRequest = new OrganizationRequest("searchquery")
                {
                    Parameters = new ParameterCollection
                    {
                        ["search"] = searchTerm,
                        ["count"] = true,
                        ["entities"] = JsonConvert.SerializeObject(entitiesArray),
                        ["options"] = JsonConvert.SerializeObject(optionsDict)
                    }
                };

                var searchResponse = orgService.Execute(searchRequest);

                if (!searchResponse.Results.Contains("response"))
                {
                    context.OutputParameters["meaf_results"] = "{\"Count\":0,\"Value\":[]}";
                    context.OutputParameters["meaf_bestscore"] = 0;
                    context.OutputParameters["meaf_bestrecord"] = null;
                    return;
                }

                string rawJson = (string)searchResponse.Results["response"];
                if (string.IsNullOrEmpty(rawJson))
                {
                    context.OutputParameters["meaf_results"] = "{\"Count\":0,\"Value\":[]}";
                    context.OutputParameters["meaf_bestscore"] = 0;
                    context.OutputParameters["meaf_bestrecord"] = null;
                    return;
                }

                var queryResults = JsonConvert.DeserializeObject<SearchQueryResults>(rawJson);
                if (queryResults == null)
                {
                    context.OutputParameters["meaf_results"] = "{\"Count\":0,\"Value\":[]}";
                    context.OutputParameters["meaf_bestscore"] = 0;
                    context.OutputParameters["meaf_bestrecord"] = null;
                    return;
                }

                tracingService.Trace($"Found {queryResults.Count} total items in fuzzy search.");

                double bestScore = 0;
                EntityReference bestRecord = null;

                if (queryResults.Value.Any())
                {
                    var bestItem = queryResults.Value.OrderByDescending(x => x.Score).First();
                    bestScore = bestItem.Score;
                    bestRecord = new EntityReference(bestItem.EntityName, new Guid(bestItem.Id));
                }

                string finalJson = JsonConvert.SerializeObject(queryResults);
                context.OutputParameters["meaf_results"] = finalJson;
                context.OutputParameters["meaf_bestscore"] = bestScore;
                context.OutputParameters["meaf_bestrecord"] = bestRecord;
            }
            catch (Exception ex)
            {
                tracingService.Trace($"Error in Plg_FuzzySearch: {ex.Message}");
                throw new InvalidPluginExecutionException($"Error in Plg_FuzzySearch: {ex.Message}", ex);
            }
        }
    }

    public class SearchQueryResults
    {
        [JsonProperty("Count")]
        public int Count { get; set; }

        [JsonProperty("Value")]
        public List<SearchResultItem> Value { get; set; } = new List<SearchResultItem>();
    }

    public class SearchResultItem
    {
        [JsonProperty("Id")]
        public string Id { get; set; }

        [JsonProperty("EntityName")]
        public string EntityName { get; set; }

        [JsonProperty("ObjectTypeCode")]
        public int ObjectTypeCode { get; set; }

        [JsonProperty("Score")]
        public double Score { get; set; }

        [JsonProperty("Attributes")]
        public Dictionary<string, object> Attributes { get; set; } = new Dictionary<string, object>();

        [JsonProperty("Highlights")]
        public Dictionary<string, List<string>> Highlights { get; set; } = new Dictionary<string, List<string>>();
    }
}

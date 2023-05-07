using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using JetBrains.Annotations;
using Newtonsoft.Json.Linq;

namespace LinkeD365.FlowToVisio
{
    public class FlowDefinition
    {
        private JObject definitionJObject;
        private string definition;
        public bool Solution { get; set; }
        public string Id { get; set; }

        [Browsable(false)]
        public string Definition
        {
            get => definition;
            set
            {
                definition = value;
                GetDefinitionJObject();
            }
        }

        [Browsable(false)]
        public JObject DefinitionJObject
        {
            get
            {
                GetDefinitionJObject();
                return definitionJObject;
            }
        }

        private void GetDefinitionJObject()
        {
            if (Category == 5 // Modern Flow
                && definitionJObject == null
               )
            {
                definitionJObject = JObject.Parse(Definition);
                ProcessDefinition(definitionJObject);
            }
        }

        private void ProcessDefinition([NotNull] JObject jObject)
        {
            if (jObject == null) throw new ArgumentNullException(nameof(jObject));

            ProcessTrigger(jObject);
            ProcessActions(jObject);
        }

        private void ProcessActions(JObject jObject)
        {
            var actions = jObject.DescendantsAndSelf().OfType<JProperty>()
                .Where(o => o.Name == "actions")
                .ToList();
            // TODO : not really the actions count - more like count of action collections
            ActionsCount = actions.Count;

            // The leaf action collections - i.e. the ones doing work :D
            var leafActions = jObject.DescendantsAndSelf().OfType<JProperty>()
                .Where(o => o.Name == "actions" 
                            // Get the leaf actions
                            && o.Descendants().OfType<JProperty>().Count(x => x.Name == "actions") == 0)
                .ToList();
            // The individual actions
            var individualLeafActions =leafActions.SelectMany(property => property.Children())
                .SelectMany(property => property.Children())
                .OfType<JProperty>()
                .ToList();

            var operationIds = individualLeafActions.Where(x => x.Descendants().OfType<JProperty>().Any(y => y.Name == "operationId"))
                .ToList();
            Operations = operationIds.Select(property =>
            {
                var operationId = property.Descendants().OfType<JProperty>()
                    .SingleOrDefault(y => y.Name == "operationId")
                    ?.Value.Value<string>();
                var entityName = property.Descendants().OfType<JProperty>()
                    .SingleOrDefault(y => y.Name == "entityName")
                    ?.Value.Value<string>();
                return new OperationAction
                {
                    OperationId = operationId,
                    EntityName = entityName
                };
            }).ToList();
        }

        public IReadOnlyList<OperationAction> Operations { get; private set; }

        public int ActionsCount { get; private set; }

        private void ProcessTrigger(JObject jObject)
        {
            var trigger = jObject.DescendantsAndSelf()
                .OfType<JProperty>()
                .FirstOrDefault(o => o.Name == "triggers");
            if (trigger != null)
            {
                var entityName = trigger.Descendants().OfType<JProperty>()
                    .FirstOrDefault(x => x.Name == "subscriptionRequest/entityname")?.Value.Value<string>();
                TriggerEntity = entityName;
                var filterExpression = trigger.Descendants().OfType<JProperty>()
                    .FirstOrDefault(x =>
                        string.Equals(x.Name, "subscriptionRequest/filterexpression", StringComparison.OrdinalIgnoreCase))
                    ?.Value.Value<string>();
                TriggerFilterExpression = filterExpression;
                var filteringAttributes = trigger.Descendants().OfType<JProperty>()
                    .FirstOrDefault(x =>
                        string.Equals(x.Name, "subscriptionRequest/filteringattributes", StringComparison.OrdinalIgnoreCase))
                    ?.Value.Value<string>();
                TriggerFilteringAttributes = filteringAttributes;
            }
        }

        public string TriggerFilteringAttributes { get; private set; }

        public string TriggerFilterExpression { get; private set; }

        [Browsable(false)]
        public bool HasTriggerEntity => !string.IsNullOrWhiteSpace(TriggerEntity);

        public string TriggerEntity { get; private set; }

        [Browsable(false)]
        [UsedImplicitly] 
        public bool HasDefinition => !string.IsNullOrWhiteSpace(Definition);

        public string Name { get; set; }
        public string Description { get; set; }
        public bool Managed { get; set; }
        public string OwnerType { get; set; }
        public bool LogicApp { get; internal set; }
        [Browsable(false)]
        public string UniqueId { get; set; }

        public string Status { get; set; }

        public string CategoryDescription
        {
            get
            {
                switch (Category)
                {
                    case 0: return $"{Category} - Workflow";
                    case 1: return $"{Category} - Dialog";
                    case 2: return $"{Category} - Business Rule";
                    case 3: return $"{Category} - Action";
                    case 4: return $"{Category} - Business Process Flow";
                    case 5: return $"{Category} - Modern Flow";
                }

                return "Unknown";
            }
        }
        [Browsable(false)]
        public int Category { get; set; }

        //public List<Comment> Comments = new List<Comment>();
    }

    public class OperationAction
    {
        public string OperationId { get; set; }
        public string EntityName { get; set; }
    }

    internal class FlowDefComparer : IComparer<FlowDefinition>
    {
        private string memberName = string.Empty; // specifies the member name to be sorted
        private SortOrder sortOrder = SortOrder.None; // Specifies the SortOrder.

        public FlowDefComparer(string strMemberName, SortOrder sortingOrder)
        {
            memberName = strMemberName;
            sortOrder = sortingOrder;
        }

        public int Compare(FlowDefinition flow1, FlowDefinition flow2)
        {
            switch (memberName)
            {
                case "Name":
                    return sortOrder == SortOrder.Ascending ? flow1.Name.CompareTo(flow2.Name) : flow2.Name.CompareTo(flow1.Name);
                case "Description":
                    return sortOrder == SortOrder.Ascending ? flow1.Description.CompareTo(flow2.Description) : flow2.Description.CompareTo(flow1.Description);
                case "Managed":
                    return sortOrder == SortOrder.Ascending ? flow1.Managed.CompareTo(flow2.Managed) : flow2.Managed.CompareTo(flow1.Managed);
                case "CategoryDescription":
                    return sortOrder == SortOrder.Ascending ? flow1.CategoryDescription.CompareTo(flow2.CategoryDescription) : flow2.CategoryDescription.CompareTo(flow1.CategoryDescription);
                case "Status":
                    return sortOrder == SortOrder.Ascending ? flow1.Status.CompareTo(flow2.Status) : flow2.Status.CompareTo(flow1.Status);
                default:
                    return sortOrder == SortOrder.Ascending ? flow1.Name.CompareTo(flow2.Name) : flow2.Name.CompareTo(flow1.Name);
            }
        }
    }
}
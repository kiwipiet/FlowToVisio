using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using JetBrains.Annotations;
using Newtonsoft.Json.Linq;

namespace LinkeD365.FlowToVisio
{
    public class FlowDefinition
    {
        private JObject definitionJObject;
        public bool Solution { get; set; }
        public string Id { get; set; }
        [Browsable(false)]
        public string Definition { get; set; }

        [Browsable(false)]
        public JObject DefinitionJObject
        {
            get
            {
                {
                    if (Category == 5 // Modern Flow
                        && definitionJObject == null
                    )
                    {
                        definitionJObject = JObject.Parse(Definition);
                    }
                }
                return definitionJObject;
            }
        }

        [UsedImplicitly] public bool HasDefinition => !string.IsNullOrWhiteSpace(Definition);

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
                    return  sortOrder == SortOrder.Ascending ? flow1.Name.CompareTo(flow2.Name) : flow2.Name.CompareTo(flow1.Name);
            }
        }
    }
}
using LinkeD365.FlowToVisio;

namespace FlowToVisioTests
{
    public class FlowDefinitionTests
    {
        [Fact]
        public void TestTriggerProcessing()
        {
            var f = new FlowDefinition
            {
                Category = 5,
                Definition = File.ReadAllText(
                    "C:\\Users\\piete\\OneDrive - DXC Production\\Documents\\ADEPT\\Documentation\\Workflows\\5 - Modern Flow\\Activated\\OnCreateUpdateExternalNotification.json")
            };
            Assert.Equal("inz_notification", f.TriggerEntity);
            Assert.Equal("statuscode", f.TriggerFilteringAttributes);
            Assert.Equal("_inz_firmmember_value eq null and _inz_quota_value eq null and _inz_variationofconditionrequestid_value eq null and statuscode eq 121570000 and _inz_externalnotificationtemplate_value ne null and (_inz_employeraccreditation_value ne null or _inz_groupvisaapplication_value ne null or _inz_jobcheck_value ne null or _inz_visaapplication_value ne null)", f.TriggerFilterExpression);
            Assert.True(f.HasTriggerEntity);
        }
        [Fact]
        public void TestActionProcessing()
        {
            var f = new FlowDefinition
            {
                Category = 5,
                Definition = File.ReadAllText(
                    "C:\\Users\\piete\\OneDrive - DXC Production\\Documents\\ADEPT\\Documentation\\Workflows\\5 - Modern Flow\\Activated\\OnCreateUpdateExternalNotification.json")
            };
            Assert.Equal("inz_notification", f.TriggerEntity);
            Assert.Equal("statuscode", f.TriggerFilteringAttributes);
            Assert.Equal("_inz_firmmember_value eq null and _inz_quota_value eq null and _inz_variationofconditionrequestid_value eq null and statuscode eq 121570000 and _inz_externalnotificationtemplate_value ne null and (_inz_employeraccreditation_value ne null or _inz_groupvisaapplication_value ne null or _inz_jobcheck_value ne null or _inz_visaapplication_value ne null)", f.TriggerFilterExpression);
            Assert.True(f.HasTriggerEntity);
        }
    }
}
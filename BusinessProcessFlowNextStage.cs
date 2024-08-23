// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.CE.Reports;
using Microsoft.Dynamics365.UIAutomation.Sample;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.CE.UCI
{
    [TestClass]
    public class BusinessProcessFlowNextStageUCI : ExtentReportParent
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly SecureString _mfaSecretKey = System.Configuration.ConfigurationManager.AppSettings["MfaSecretKey"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        [TestCategory("BPF")]
        // [Priority(0)]
        public void UCITestBusinessProcessFlowNextStage()
        {
            var client = new WebClient(TestSettings.Options);
            using (var xrmApp = new XrmApp(client))

                try
                {


                    xrmApp.OnlineLogin.Login(_xrmUri, _username, _password, _mfaSecretKey);

                    xrmApp.Navigation.OpenApp(UCIAppName.Sales);

                    xrmApp.Navigation.OpenSubArea("Sales", "Leads");

                    //open first lead
                    xrmApp.ThinkTime(5000);
                    xrmApp.Grid.OpenRecord(0);

                    //move to Qualify stage
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SelectStage("Qualify");
                    xrmApp.ThinkTime(3000);

                    //Open and Entering the Qulify stage opetion values... 
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("parentcontactid", "Hazari Arshi");
                    xrmApp.Lookup.OpenRecord(0);


                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("parentaccountid", "Hazari Tech");
                    xrmApp.Lookup.OpenRecord(0);

                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue(new OptionSet { Name = "header_process_purchasetimeframe", Value = "This Quarter" });
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("budgetamount", "500000");
                    xrmApp.ThinkTime(5000);


                    xrmApp.BusinessProcessFlow.SetValue("decisionmaker");
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("description", "Testing Description");

                    //closing Qualify stage
                    xrmApp.ThinkTime(3000);
                    xrmApp.BusinessProcessFlow.Close("Qualify");

                    //qualify the lead via click button
                    xrmApp.ThinkTime(3000);
                    xrmApp.CommandBar.ClickCommand("Qualify");
                    xrmApp.ThinkTime(5000);

                    //moving to the BPF next stage(Develop)...
                    xrmApp.BusinessProcessFlow.SelectStage("Develop");
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("identifycompetitors");
                    xrmApp.ThinkTime(5000);


                    //moving to the BPF next stage
                    xrmApp.BusinessProcessFlow.NextStage("Develop");
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.NextStage("Propose");
                    xrmApp.ThinkTime(5000);
                    xrmApp.BusinessProcessFlow.SetValue("completefinalproposal");
                    xrmApp.ThinkTime(5000);

                    //Close and Lost... with required data... 
                    xrmApp.Dialogs.CloseOpportunity(true, 600.0855, new DateTime(), "Closing opportunity!");
                    xrmApp.ThinkTime(5000);


                }
                catch (Exception e)
                {

                    TakeScreenshot(client, "UCITestBusinessProcessFlowNextStage" + CaseCount);
                    getErrorLogs(client);
                    markTestcaseSatus(e);
                }

        }
    }
}

using kCura.Relativity.Client;
using kCura.Relativity.Client.DTOs;
using NUnit.Framework;
using Relativity.API;
using Relativity.Services.Objects;
using Relativity.Services.Objects.DataContracts;
using Relativity.Test.Helpers.ServiceFactory.Extentions;
using Relativity.Test.Helpers.SharedTestHelpers;
using Relativity.Test.Helpers.WorkspaceHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using QueryResult = kCura.Relativity.Client.QueryResult;

namespace delimiter_count_by_saved_search
{
    [TestFixture]
    public class ScriptDelimiterCountBySavedSearchTest
    {

        #region variables
        private const string _NSERIO_OBJECTTYPE_NAME = "NSerio";

        private IRSAPIClient _client;
        private IObjectManager _objectManagerClient;
        private int _workspaceId;
        private int _artifactTypeID;
        private readonly string _workspaceName = ConfigurationHelper.TEST_WORKSPACE_NAME;
        private readonly string _workspaceTemplateName = ConfigurationHelper.TEST_WORKSPACE_TEMPLATE_NAME;
        private IServicesMgr _servicesManager;
        private const string _SCRIPT_NAME = "Delimiter Count by Saved Search";
        private const string _SAVED_SEARCH_NAME = "Diana Test";
        private const string _SOURCE_FIELD = "EmailTo";
        private const string _DELIMITER = ";";
        private const string _COUNT_DESTINATION_FIELD = "Count";
        private bool _workspaceCreatedByTest;

        #endregion

        #region Setup

        [OneTimeSetUp]
        public void Execute_TestFixtureSetup()
        {
            //Setup for testing
            //Create a new instance of the Test Helper

            var helper = Relativity.Test.Helpers.TestHelper.System();
            _servicesManager = helper.GetServicesManager();

            //create client
            _client = helper.GetServicesManager().GetProxy<IRSAPIClient>(ConfigurationHelper.ADMIN_USERNAME, ConfigurationHelper.DEFAULT_PASSWORD);
            _objectManagerClient = helper.GetServicesManager().GetProxy<IObjectManager>(ConfigurationHelper.ADMIN_USERNAME, ConfigurationHelper.DEFAULT_PASSWORD);
            //Get workspace ID of the workspace for Nserio or Create a workspace
            _workspaceId = GetWorkspaceId(_workspaceName, _objectManagerClient);
            if (_workspaceId == 0) //-- if no workspace found, create it
            {
                _workspaceId = CreateWorkspace.Create(_client, _workspaceName, _workspaceTemplateName);
                _workspaceCreatedByTest = true;
            }

            _client.APIOptions.WorkspaceID = _workspaceId;

            var path = GetLocalDocumentsFolderPath();

            //Import Application to the workspace
            //File path of the Test App
            string[] path1 = { path, "RA_Delimiter-Count-By-Saved-Search-Test-APP.rap" };
            string filepathTestApp = Path.Combine(path1);

            //File path of the application containing the actual script
            string[] path2 = { path, "RA_Delimiter-Count-By-Saved-Search.rap" };
            string filepathApp = Path.Combine(path2);


            //Importing the applications
            Relativity.Test.Helpers.Application.ApplicationHelpers.ImportApplication(_client, _workspaceId, true, filepathTestApp);
            Relativity.Test.Helpers.Application.ApplicationHelpers.ImportApplication(_client, _workspaceId, true, filepathApp);

            //set artifacttypeid
            _artifactTypeID = (int)ArtifactType.Document;

            //Import Documents to workspace
            ImportHelper.Import.ImportDocument(_workspaceId, path);
        }

        #endregion


        #region Teardown
        [OneTimeTearDown]
        public void Execute_TestFixtureTeardown()
        {
            if (_workspaceCreatedByTest)
            { //-- delete the workspace created by the test execution
                DeleteWorkspace.Delete(_client, _workspaceId);
            }
        }

        #endregion

        #region Tests
        [Test]
        [Description("Verify the Relativity Script executes succesfully")]
        public void Integration_Test_Golden_Flow_Valid()
        {
            //Act
            var scriptResults = ExecuteScript_DelimiterCountBySavedSearch(_SCRIPT_NAME, _SAVED_SEARCH_NAME, _SOURCE_FIELD, _DELIMITER, _COUNT_DESTINATION_FIELD);

            //Assert
            Assert.AreEqual(true, scriptResults.Success);
        }

        [Test]
        [Description("Verify object records are created successfully")]
        public void Integration_Test_Check_Created_Records()
        {
            //Act
            var scriptResults = ExecuteScript_DelimiterCountBySavedSearch(_SCRIPT_NAME, _SAVED_SEARCH_NAME, _SOURCE_FIELD, _DELIMITER, _COUNT_DESTINATION_FIELD);
            var objectsWhereCreatedSuccessfully = GetCreatedObjectsStatus();
            //Assert
            Assert.AreEqual(true, scriptResults.Success);
            Assert.AreEqual(true, objectsWhereCreatedSuccessfully);
            
        }

        #endregion

        #region Helpers

        private string GetLocalDocumentsFolderPath()
        {
            //string path = Path.Combine(Environment.CurrentDirectory, "Resources");
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            string binFolderPath = Path.GetDirectoryName(executingAssembly.Location);
            if (binFolderPath == null)
            {
                throw new Exception("Bin folder path is empty");
            }

            string path = Path.Combine(binFolderPath);
            return path;
        }
        public bool GetCreatedObjectsStatus()
        {
            bool state = false;
            QueryRequest request = new QueryRequest
            {
                ObjectType = new ObjectTypeRef { ArtifactTypeID = _artifactTypeID },
                Fields = new[]
                    {
                        new FieldRef { Name = _COUNT_DESTINATION_FIELD }
                    }
            };

            Task<QueryResultSlim> taskForResult = _objectManagerClient.QuerySlimAsync(_workspaceId, request, 0, int.MaxValue);
            QueryResultSlim result = GetQueryResultFromTask(taskForResult);

            if (result.TotalCount > 0)
            {
                string extractedinfo = result.Objects[0].Values[0].ToString();
                if (extractedinfo.Contains("2"))
                {
                    state = true;
                }
            }

            return state;
        }

        public QueryResultSlim GetQueryResultFromTask(Task<QueryResultSlim> task)
        {
            QueryResultSlim result = task.ConfigureAwait(false).GetAwaiter().GetResult();
            return result;
        }

        public RelativityScriptResult ExecuteScript_DelimiterCountBySavedSearch(string scriptName, string savedSearchName, string sourceFieldName, string delimiter, string countDestinationFieldName)
        {
            _client.APIOptions.WorkspaceID = _workspaceId;

            //Retrieve script by name
            Query<RelativityScript> relScriptQuery = new Query<RelativityScript>
            {
                Condition = new TextCondition(RelativityScriptFieldNames.Name, TextConditionEnum.EqualTo, scriptName),
                Fields = FieldValue.AllFields
            };

            QueryResultSet<RelativityScript> relScriptQueryResults = _client.Repositories.RelativityScript.Query(relScriptQuery);
            if (!relScriptQueryResults.Success)
            {
                throw new Exception(String.Format("An error occurred finding the script: {0}", relScriptQueryResults.Message));
            }

            if (!relScriptQueryResults.Results.Any())
            {
                throw new Exception(String.Format("No results returned: {0}", relScriptQueryResults.Message));
            }

            //Retrieve script inputs
            RelativityScript script = relScriptQueryResults.Results[0].Artifact;
            var inputnames = GetRelativityScriptInput(_client, scriptName);
            int savedsearchartifactid = Query_For_Saved_SearchID(savedSearchName, _client);

            //Set inputs for script
            RelativityScriptInput input = new RelativityScriptInput(inputnames[0], savedsearchartifactid.ToString());
            RelativityScriptInput input2 = new RelativityScriptInput(inputnames[1], sourceFieldName);
            RelativityScriptInput input3 = new RelativityScriptInput(inputnames[2], delimiter);
            RelativityScriptInput input4 = new RelativityScriptInput(inputnames[3], countDestinationFieldName);

            //Execute the script
            List<RelativityScriptInput> inputList = new List<RelativityScriptInput> { input, input2, input3, input4 };

            RelativityScriptResult scriptResult = null;

            try
            {
                scriptResult = _client.Repositories.RelativityScript.ExecuteRelativityScript(script, inputList);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: {0}", ex.Message);
            }

            //Check for success.
            if (!scriptResult.Success)
            {
                Console.WriteLine(string.Format(scriptResult.Message));
            }
            else
            {
                int observedOutput = scriptResult.Count;
                Console.WriteLine("Result returned: {0}", observedOutput);

            }

            return scriptResult;
        }

        public static List<string> GetRelativityScriptInput(IRSAPIClient client, string scriptName)
        {
            var returnval = new List<string>();
            List<RelativityScriptInputDetails> scriptInputList;

            int artifactid = GetScriptArtifactId(scriptName, client);

            // STEP 1: Using ArtifactID, set the script you want to run.
            RelativityScript script = new RelativityScript(artifactid);

            // STEP 2: Call GetRelativityScriptInputs.
            try
            {
                scriptInputList = client.Repositories.RelativityScript.GetRelativityScriptInputs(script);
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("An error occurred: {0}", ex.Message));
                return returnval;
            }

            // STEP 3: Each RelativityScriptInputDetails object can be used to generate a RelativityScriptInput object, 
            // but this example only displays information about each input.
            foreach (RelativityScriptInputDetails relativityScriptInputDetails in scriptInputList)
            {
                returnval.Add(relativityScriptInputDetails.Name);
            }
            return returnval;
        }

        public static int GetScriptArtifactId(string scriptName, IRSAPIClient _client)
        {
            int ScriptArtifactId = 0;
            try
            {
                Query newQuery = new Query();
                TextCondition queryCondition = new TextCondition(RelativityScriptFieldNames.Name, TextConditionEnum.Like, scriptName);
                newQuery.Condition = queryCondition;
                newQuery.ArtifactTypeID = 28;
                _client.APIOptions.StrictMode = false;
                var results = _client.Query(_client.APIOptions, newQuery);
                ScriptArtifactId = results.QueryArtifacts[0].ArtifactID;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: {0}", ex.Message);
            }

            return ScriptArtifactId;
        }

        public static int GetWorkspaceId(string workspaceName, IObjectManager _client)
        {
            int workspaceArtifactId = 0;           

            try
            {
                QueryRequest queryRequest = new QueryRequest();
                queryRequest.ObjectType = new ObjectTypeRef { ArtifactTypeID = (int)ArtifactType.Case };
                queryRequest.Condition = $"('{WorkspaceFieldNames.Name}' IN ['{workspaceName}'])";
                QueryResultSlim results = _client.QuerySlimAsync(-1, queryRequest, 1, 1)
                    .ConfigureAwait(false)
                    .GetAwaiter()
                    .GetResult();
                workspaceArtifactId = results.Objects[0].ArtifactID;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: {0}", ex.Message);
            }

            return workspaceArtifactId;
        }

        public static int Query_For_Saved_SearchID(string savedSearchName, IRSAPIClient _client)
        {

            int searchArtifactId = 0;

            var query = new Query
            {
                ArtifactTypeID = (int)ArtifactType.Search,
                Condition = new TextCondition("Name", TextConditionEnum.Like, savedSearchName)
            };
            QueryResult result = null;

            try
            {
                result = _client.Query(_client.APIOptions, query);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: {0}", ex.Message);
            }

            if (result != null)
            {
                searchArtifactId = result.QueryArtifacts[0].ArtifactID;
            }

            return searchArtifactId;
        }

        public static bool DeleteAllObjectsOfSpecificTypeInWorkspace(IRSAPIClient proxy, int workspaceID, int artifactTypeID)
        {
            proxy.APIOptions.WorkspaceID = workspaceID;

            //Query RDO
            WholeNumberCondition condition = new WholeNumberCondition("Artifact ID", NumericConditionEnum.IsSet);

            kCura.Relativity.Client.DTOs.Query<RDO> query = new Query<RDO>
            {
                ArtifactTypeID = artifactTypeID,
                Condition = condition
            };

            QueryResultSet<RDO> results = new QueryResultSet<RDO>();
            results = proxy.Repositories.RDO.Query(query);

            if (!results.Success)
            {
                Console.WriteLine("Error deleting the object: " + results.Message);

                for (Int32 i = 0; i <= results.Results.Count - 1; i++)
                {
                    if (!results.Results[i].Success)
                    {
                        proxy.Repositories.RDO.Delete(results.Results[i].Artifact);
                    }
                }
            }

            return true;
        }

        #endregion

    }

}

﻿using kCura.Relativity.DataReaderClient;
using kCura.Relativity.ImportAPI;
using Relativity.Test.Helpers.SharedTestHelpers;
using System;
using System.Data;
using System.IO;

namespace delimiter_count_by_saved_search.ImportHelper
{
    public class Import
    {
        private static readonly string IMPORT_API_ENDPOINT = $"{ConfigurationHelper.SERVER_BINDING_TYPE}://{ConfigurationHelper.REST_SERVER_ADDRESS}/Relativitywebapi/";

        public static void ImportDocument(int workspaceId, string path)
        {
            int identifyFieldArtifactID = 1003667;    // 'Control Number' Field

            ImportAPI iapi = new ImportAPI(ConfigurationHelper.ADMIN_USERNAME, ConfigurationHelper.DEFAULT_PASSWORD, IMPORT_API_ENDPOINT);

            var importJob = iapi.NewNativeDocumentImportJob();

            importJob.OnMessage += ImportJobOnMessage;
            importJob.OnComplete += ImportJobOnComplete;
            importJob.OnFatalException += ImportJobOnFatalException;
            importJob.Settings.CaseArtifactId = workspaceId;
            importJob.Settings.ExtractedTextFieldContainsFilePath = false;

            importJob.Settings.NativeFilePathSourceFieldName = "Original Folder Path";
            importJob.Settings.NativeFileCopyMode = NativeFileCopyModeEnum.CopyFiles; // NativeFileCopyModeEnum.CopyFiles; NativeFileCopyModeEnum.DoNotImportNativeFiles
            importJob.Settings.OverwriteMode = OverwriteModeEnum.AppendOverlay;

            importJob.Settings.IdentityFieldId = identifyFieldArtifactID;
            importJob.SourceData.SourceData = GetDocumentDataTable(path).CreateDataReader();

            Console.WriteLine("=======>>>>> Executing import...");

            importJob.Execute();

        }

        public static DataTable GetDocumentDataTable(string path)
        {
            DataTable table = new DataTable();

            // The document identifer column name must match the field name in the workspace.
            table.Columns.Add("Control Number", typeof(string));
            table.Columns.Add("Original Folder Path", typeof(string));
            table.Columns.Add("Email From", typeof(string));
            table.Columns.Add("Email To", typeof(string));
            table.Columns.Add("Email CC", typeof(string));
            table.Columns.Add("Email BCC", typeof(string));
            //string filePath = @"\SampleFile.txt";
            string[] paths = { path, @"SampleFile.txt" };
            string filePath = Path.Combine(paths);
            string content = File.ReadAllText(filePath);
            table.Rows.Add("11Works", filePath, content, content, content, content);

            return table;
        }

        static void ImportJobOnMessage(Status status)
        {
            Console.WriteLine("Message: {0}", status.Message);
        }

        static void ImportJobOnFatalException(JobReport jobReport)
        {
            Console.WriteLine("Fatal Error: {0}", jobReport.FatalException);
        }

        static void ImportJobOnComplete(JobReport jobReport)
        {
            Console.WriteLine("Job Finished With {0} Errors: ", jobReport.ErrorRowCount);
        }

    }
}


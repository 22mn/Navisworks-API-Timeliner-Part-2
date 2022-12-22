using System;
using System.IO;
using System.Globalization;
using Microsoft.Win32;
using System.Diagnostics;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Timeliner;

namespace Timeliner_Part2
{
    // name, developer id & displayname
    [PluginAttribute("CreateTaskFromDataSource", "TwentyTwo",
        DisplayName = "Timeliner Part-2 DataSource")]
    // location of an AddInPlugin in the Navisworks gui 
    [AddInPlugin(AddInLocation.None)]
    // plugin group - timeliner data source provider
    [Interface("TimelinerDataSourceProvider", "Navisworks",
        DisplayName = "Timeliner Part-2 DataSource")]

    public sealed class MainClass : TimelinerDataSourceProvider
    {
        public MainClass() { }

        // called to create data source
        public override TimelinerDataSource CreateDataSource(string displayName)
        {
            TimelinerDataSource dataSource = new TimelinerDataSource(displayName)
            {
                //identify datasource
                ProjectIdentifier = @"D:\TwentyTwo\Navisworks_Tutorial\Timeliner Part2 Resources\data-source.csv",
                DataSourceProviderId = base.Id,
                DataSourceProviderVersion = 1.0,
                DataSourceProviderName = "CreateTaskFromDataSource"
            };
            // add external fields to map
            AddAvailableFields(dataSource);

            return dataSource;           
        }

        // called when the user selects Rebuild or Synchronize
        protected override TimelinerImportTasksResult ImportTasksCore(TimelinerDataSource dataSource)
        {

            // new timelinerimporttaskresult
            TimelinerImportTasksResult importResult = new TimelinerImportTasksResult();

            // set true the data source was configured to set the SimulationTaskTypeName
            importResult.TaskTypeWasSet = true;

            // create root task
            TimelinerTask rootTask = new TimelinerTask();

            // to associate this task with an external task
            rootTask.SynchronizationId = TimelinerTask.DataSourceRootTaskIdentifier;

            // read data source - csv file 
            StreamReader fileData = new StreamReader(dataSource.ProjectIdentifier);

            // culture info
            CultureInfo cultureInfo = new CultureInfo("en-US");

            // read first line;
            string strData = fileData.ReadLine();
            
            
            while (!fileData.EndOfStream)
            {
                // read data - store in fields
                strData = fileData.ReadLine();
                string[] fields = strData.Split(',');
                
                // create sub-task
                TimelinerTask task = new TimelinerTask()
                {
                    // assign field values
                    SynchronizationId = fields[0],
                    DisplayName = fields[1],
                    PlannedStartDate = StringToDateTime(fields[6]),
                    PlannedEndDate = StringToDateTime(fields[7]),
                    SimulationTaskTypeName = fields[8]
                };
                // attach selectionsets for simulation
                task.Selection.CopyFrom(getSelectionSetByName(fields[1]));
                // add to root task
                rootTask.Children.Add(task);
            }
            // close csv data file
            fileData.Close();
            // set importresult root task
            importResult.RootTask = rootTask;

            return importResult;
        }

        // called when the user clicks edit datasoruce 
        public override void UpdateDataSource(TimelinerDataSource dataSource)
        {
            // check provided datasource file
            FileReferenceResolver resolver = new FileReferenceResolver();
            FileReferenceResolveResult result = resolver.Resolve(dataSource.ProjectIdentifier);
            if (result.Response == FileResolutionResponse.OK)
            {
                dataSource.ProjectIdentifier = result.FileNameToOpen;
            }
            // fields mapping
            if (dataSource.AvailableFields.Count == 0)
            {
                AddAvailableFields(dataSource);
            }
        }
        // add external fields to map timeliner fields
        public void AddAvailableFields(TimelinerDataSource dataSource)
        {
            // fields add to datasource
            dataSource.AvailableFields.Add(new TimelinerDataSourceField("UniqueID", "UniqueId"));
            dataSource.AvailableFields.Add(new TimelinerDataSourceField("DisplayName", "Display Name"));
            dataSource.AvailableFields.Add(new TimelinerDataSourceField("PlannedStart", "Planned Start"));
            dataSource.AvailableFields.Add(new TimelinerDataSourceField("PlannedEnd", "Planned End"));
            dataSource.AvailableFields.Add(new TimelinerDataSourceField("TaskType", "Task Type"));

        }

        // get selectionset by name
        public SelectionSourceCollection getSelectionSetByName(string name)
        {
            // current document
            Document doc = Application.ActiveDocument;
            // collect selection sets
            SavedItemCollection selectionSets = doc.SelectionSets.RootItem.Children;
            foreach (SavedItem set in selectionSets)
            {
                if (set.DisplayName == name)
                {
                    // saveditem to selection set  
                    SelectionSet selSect = set as SelectionSet;
                    // create selection source from selection set
                    SelectionSource selSource = doc.SelectionSets.CreateSelectionSource(selSect);
                    // create selection source collection from selection source
                    SelectionSourceCollection selSourceCol = new SelectionSourceCollection() { selSource };
                    return selSourceCol;
                }

            }
            // return empty source
            return new SelectionSourceCollection();
        }

        // text convert to datetime
        public DateTime StringToDateTime(string str)
        {
            string[] strList = str.Split('/');
            int day = Int32.Parse(strList[0]);
            int month = Int32.Parse(strList[1]);
            int year = Int32.Parse(strList[2]);

            DateTime date = new DateTime(year, month, day);

            return date;
        }
        
        // inherited method implementation
        public override TimelinerValidateSettingsResult ValidateSettings(TimelinerDataSource dataSource)
        {
            return new TimelinerValidateSettingsResult(true);
        }

        // inherited method implementation
        protected override void DisposeManagedResources() {}

        // inherited method implementation
        protected override void DisposeUnmanagedResources() {}

        // inherited method implementation
        public override bool IsAvailable => true;
    }
}

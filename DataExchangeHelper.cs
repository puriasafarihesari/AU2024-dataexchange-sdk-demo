using Autodesk.DataExchange;
using Autodesk.DataExchange.Core.Enums;
using Autodesk.DataExchange.DataModels;
using Autodesk.DataExchange.SchemaObjects.Units;
using Autodesk.GeometryPrimitives.Design;
using Autodesk.GeometryPrimitives.Geometry;
using Autodesk.GeometryPrimitives.Math;
using Autodesk.DataExchange.Models;
using Autodesk.Parameters;
using Autodesk.DataExchange.Core.Models;
using System.Threading.Tasks;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Collections;
using System.Data;

namespace AU2024_smart_parameter_updater
{
    internal class DataExchangeHelper
    {
        private Autodesk.DataExchange.Client _client = null;
        private string _hubId;
        private string _projectId;
        private string _folderUrn;
        private string _collectionId;

        public void Connect(DemoConfiguration cfg)
        {
            try
            {
                //connect
                if (_client == null)
                {
                    _client = new Autodesk.DataExchange.Client(new SDKOptionsDefaultSetup()
                    {
                        ApplicationName = cfg.ApplicationName,
                        ClientId = cfg.ClientId,
                        ClientSecret = cfg.ClientSecret,
                        CallBack = cfg.CallBack
                    });

                    //creates a log
                    _client.SDKOptions.Logger.SetDebugLogLevel();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            };
        }

        public void SetFolder(string hubId, string projectId, string folderUrn, string collectionId)
        {
            this._hubId = hubId;
            this._projectId = projectId;
            this._folderUrn = folderUrn;
            this._collectionId = collectionId;
        }


        public async Task<ExchangeData> ReadDataExchange(string exchangeFileUrn, string category)
        {
            try
            {
                ExchangeDetails exchangeDetails = await _client.GetExchangeDetailsAsync(exchangeFileUrn).ConfigureAwait(false);

                //Workaround for v4.0.4 to get the hubId
                List<ProjectInfo> allHubsProjects = await _client.SDKOptions.HostingProvider.GetHubsProjectsAsync().ConfigureAwait(false);
                string hubId = allHubsProjects.Where(item => item.ProjectId == exchangeDetails.ProjectUrn).Select(item => item.HubId).FirstOrDefault();

                //1 - define what exchange 
                DataExchangeIdentifier exchangeIdentifier = new DataExchangeIdentifier
                {
                    CollectionId = exchangeDetails.CollectionID,
                    ExchangeId = exchangeDetails.ExchangeID,
                    HubId = hubId
                };

                //2 - retrieves the data from the specified exchange
                ExchangeData exchangeData = await _client.GetExchangeDataAsync(exchangeIdentifier);

                //3 - an ElementDataModel object is created with the data from ACC 
                ElementDataModel elDataModel = ElementDataModel.Create(_client, exchangeData);

                //4 - we can now loop on the data and see what is inside 
                IEnumerable<Element> wallElements = elDataModel.Elements.Where(item => item.Category == category);

                foreach (Element wallElement in wallElements)
                {
                    if (wallElement.InstanceParameters != null)
                    {
                        foreach (Autodesk.DataExchange.DataModels.Parameter instanceParameter in wallElement.InstanceParameters)
                        {
                            Console.WriteLine($"name: {instanceParameter.Name}, value: {instanceParameter.Value}, type: {instanceParameter.Description}");                            
                        }
                    }
                }

                return exchangeData;
            }
            catch (Exception ex) {
                throw ex;                
            }
        }

        public async Task AddElementsToExchange(ExchangeData exchangeData, ExchangeDetails exchangeDetails, string category, DataTable table)
        {
            var ifcGuid = "";
            var elDataModel = ElementDataModel.Create(_client, exchangeData);
            var wallElements = elDataModel.Elements.Where(item => item.Category == category);
            var schemaIdBaseString = "exchange.parameter." + exchangeDetails.SchemaNamespace + ":String";

            try
            {
                ElementDataModel model = ElementDataModel.Create(_client);

                foreach (Element wallElement in wallElements)
                {
                    List<ElementGeometry> wall1Geometries = new List<ElementGeometry>();
                    List<ElementGeometry> elGeometries = await elDataModel.GetElementGeometryByElementAsync(wallElement);
                    
                    //Each element can have a list of geometries attached to it.                        
                    foreach (ElementGeometry elGeometry in elGeometries)
                    {
                        Autodesk.DataExchange.DataModels.Geometry g = elGeometry as Autodesk.DataExchange.DataModels.Geometry;
                        Autodesk.DataExchange.DataModels.Geometry wallGeometry = ElementDataModel.CreateGeometry(new GeometryProperties(g.FilePath, g.RenderStyle));
                        Console.WriteLine("Copying the geometry " + g.FilePath);
                        wall1Geometries.Add(wallGeometry);
                    }

                    Element wall = model.AddElement(
                        new ElementProperties
                        {
                            ElementId = wallElement.Id, Name = wallElement.Name, Category = wallElement.Category
                        }
                    );

                    foreach (Autodesk.DataExchange.DataModels.Parameter param in wallElement.InstanceParameters)
                    {
                        ParameterDefinition pDef = null;

                        if (param.ParameterDataType == ParameterDataType.String)
                        {                            
                            pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.String);
                            (pDef as StringParameterDefinition).Value = param.Value;
                        }

                        else if (param.ParameterDataType == ParameterDataType.Float64)
                        {
                            pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.Float64);
                            (pDef as MeasurableParameterDefinition).Value = param.Value;
                        }

                        else if (param.ParameterDataType == ParameterDataType.Int64)
                        {
                            pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.Int64);
                            (pDef as Int64ParameterDefinition).Value = param.Value;
                        }

                        else if (param.ParameterDataType == ParameterDataType.Int32)
                        {
                            if (param.SpecName == "Enumeration") // this may not be the best way to handle enums
                            {
                                pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.String);
                                (pDef as StringParameterDefinition).Value = param.Value;
                            }
                            else
                            {
                                pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.Int32);
                                (pDef as Int32ParameterDefinition).Value = param.Value;
                            }
                        }

                        else if (param.ParameterDataType == ParameterDataType.Bool)
                        {                            
                            pDef = ParameterDefinition.Create(param.SchemaId, ParameterDataType.Bool);
                            (pDef as BoolParameterDefinition).Value = param.Value;
                        }

                        if (pDef != null)
                        {
                            Console.WriteLine($"Copying param {param.Name} - {param.Value} ");
                            await wall.CreateParameterAsync(pDef);
                        }

                        if (param.Name == "IfcGUID")
                        {
                            ifcGuid = param.Value;
                        }
                    }

                    Dictionary<string, string> dataFromExcel = ExcelHelper.GetDataFromExcel(table, ifcGuid);

                    if (dataFromExcel.Count > 0)
                    {
                        foreach (KeyValuePair<string, string> entry in dataFromExcel)
                        {
                            string schemaId = schemaIdBaseString + entry.Key + "-1.0.0";
                            Console.WriteLine($"Adding new param {entry.Key} - {entry.Value} ");
                            ParameterDefinition customParam = ParameterDefinition.Create(schemaId, ParameterDataType.String);
                            customParam.Name = entry.Key;
                            customParam.SampleText = "sample text";
                            customParam.Description = "this i a description";
                            customParam.ReadOnly = false;
                            customParam.GroupID = Group.General.DisplayName();
                            (customParam as StringParameterDefinition).Value = entry.Value;
                            await wall.CreateParameterAsync(customParam);
                        }
                    }
                    model.SetElementGeometryByElement(wall, wall1Geometries);
                 }

                DataExchangeIdentifier exchangeIdentifier = new DataExchangeIdentifier
                {
                    CollectionId = exchangeDetails.CollectionID,
                    ExchangeId = exchangeDetails.ExchangeID,
                    HubId = exchangeDetails.HubId,
                };

                //use the ExchangeData object exposed by the ElementModel to sync with the cloud
                await _client.SyncExchangeDataAsync(exchangeIdentifier, model.ExchangeData).ConfigureAwait(false);

                //create the ACC viewable 
                await _client.GenerateViewableAsync(exchangeDetails.ExchangeID, exchangeDetails.CollectionID).ConfigureAwait(false);
                
                //show the revisions collection data
                IEnumerable<ExchangeRevision> revisions = await _client.GetExchangeRevisionsAsync(exchangeIdentifier);
                foreach (ExchangeRevision revision in revisions)
                {
                    Console.WriteLine($"Revision in exchange {exchangeIdentifier},  {revision.Id}");                    
                }
            }
            catch (Exception ex)
            {
                throw ex;
            };
        }


        /// <summary>
        /// This function creates a container for the exchange in ACC
        /// </summary>
        public async Task<ExchangeDetails> CreateExchange(string exchangeName, string description)
        {
            try 
            { 
                string name = exchangeName;            
                ProjectInfo projectDetails = await _client.SDKOptions.HostingProvider.GetProjectInformationAsync(this._hubId, this._projectId);
                ProjectType projectType = ProjectType.ACC;

                //1 - define a request object to create the exchange in ACC
                ExchangeCreateRequestACC exchangeCreateRequest = new ExchangeCreateRequestACC()
                {
                    Host = _client.SDKOptions.HostingProvider,
                    Contract = new Autodesk.DataExchange.ContractProvider.ContractProvider(),
                    Description = description,
                    FileName = name,
                    ACCFolderURN = this._folderUrn,
                    ProjectId = projectDetails.ProjectId,
                    Region = projectDetails.HubRegion,
                    HubId = projectDetails.HubId,
                    CollectionId = this._collectionId,
                    ProjectType = projectType,
                };
                
                return await _client.CreateExchangeAsync(exchangeCreateRequest);
            }
            catch (Exception ex)
            {
                throw ex;
            };
        }

    }
}


namespace AD_Dataverse_Framework
{
    using AD_Dataverse_Framework.Helper;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Tooling.Connector;
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// This class contains the operation which can be performed on the dataverse.
    /// </summary>
    public static class DataverseOperation_WebApi
    {
        /// <summary>
        /// Azure Application Client Id.
        /// </summary>
        public static string applicationUserClientId;

        /// <summary>
        /// Azure Application Secret Key.
        /// </summary>
        public static string applicationUserClientSecret;

        /// <summary>
        /// Dataverse URL where the operations needs to be performed.
        /// </summary>
        public static string dataverseURL;

        /// <summary>
        /// This method would validate the connection objects set an return the CrmServiceClient object.
        /// </summary>
        /// <returns>Object of CrmServiceClient</returns>
        public static CrmServiceClient GetCrmServiceClient()
        {
            try
            {
                var conn = new CrmServiceClient($@"AuthType=ClientSecret;url={dataverseURL};ClientId={applicationUserClientId};ClientSecret={applicationUserClientSecret}");

                return conn;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// This method would create the record in dataverse.
        /// </summary>
        /// <param name="entity">Entity object of the Entity for which records has to be created.</param>
        /// <returns cref="Guid" >Return the record guid of the entity record created.</returns>
        public static Guid CreateRecord(Entity entity)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady && entity != null)
                {
                    Guid RecordID = crmServiceClient.Create(entity);
                    return RecordID; ;
                }
                return Guid.Empty;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// This method would update the record in dataverse.
        /// </summary>
        /// <param name="entity">Entity object of the Entity for which records has to be updated.</param>
        public static void UpdateRecord(Entity entity)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady && entity != null)
                {
                    crmServiceClient.Update(entity);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// This method would associate the records with the base entity record
        /// </summary>
        /// <param name="entityName">Entity logical name.</param>
        /// <param name="entityId">Entity record id (guid) for which the relationship record has to be set.</param>
        /// <param name="relationshipName">Relationship field logical name.</param>
        /// <param name="destinationEntity">Entity record for which the base entity is to be assocaited with.</param>
        /// <param name="destinationRecordGuid">Entity record id (guid) of destination entity reocrd for which the relationship record has to be set.</param>
        public static void Associate(string entityName, Guid entityId, string relationshipName, Entity destinationEntity, Guid destinationRecordGuid)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady &&
                    !string.IsNullOrEmpty(entityName) && entityId != Guid.Empty &&
                    !string.IsNullOrEmpty(relationshipName) && destinationEntity != null && destinationRecordGuid != Guid.Empty)
                {
                    crmServiceClient.Associate(entityName,
                        entityId,
                        new Relationship(relationshipName),
                        new EntityReferenceCollection() { new EntityReference(destinationEntity.LogicalName, destinationRecordGuid) }
                        );
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// This method will retrieve single record from dataverse.
        /// </summary>
        /// <param name="entityName">Entity logical name for which record is to be retrieved.</param>
        /// <param name="recordGuid">Entity record id (guid) for which record is to retieved.</param>
        /// <param name="columnList">Array of entity fields to be retrieved.</param>
        /// <returns cref="Entity" >The Entity record.</returns>
        public static Entity RetrieveSingleRecord(string entityName, Guid recordGuid, string[] columnList)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady
                    && !string.IsNullOrEmpty(entityName) && recordGuid != Guid.Empty && columnList != null && columnList.Any())
                {
                    var columnsToRetrieve = new ColumnSet();
                    columnsToRetrieve.AddColumns(columnList);
                    Entity entity = crmServiceClient.Retrieve(entityName, recordGuid, columnsToRetrieve);
                    return entity;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// This method will retrieve multiple records from dataverse.
        /// </summary>
        /// <param name="entityName">Entity logical name for which record is to be retrieved.</param>
        /// <param name="filter">Filter expression to filter data.</param>
        /// <param name="columnList">Array of entity fields to be retrieved.</param>
        /// <returns cref="EntityCollection">The Entity Collection.</returns>
        public static EntityCollection RetrieveMultipleRecordFromCRM(string entityName, FilterExpression filter, string[] columnList)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady
                    && !string.IsNullOrEmpty(entityName) && filter != null && columnList != null && columnList.Any())
                {
                    var columnsToRetrieve = new ColumnSet();
                    columnsToRetrieve.AddColumns(columnList);

                    QueryExpression qe = new QueryExpression(entityName)
                    {
                        Criteria = filter,
                        ColumnSet = columnsToRetrieve
                    };

                    EntityCollection ec = crmServiceClient.RetrieveMultiple(qe);
                    return ec;
                }
                return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// This method will perform bulk operations on the entity data list
        /// </summary>
        /// <param name="bulkAction">Enum value for bulk operation.</param>
        /// <param name="entityDataList">Collection of Entity object.</param>
        /// <returns cref="ExecuteMultipleResponse">Collection of Response for multiple request</returns>
        public static ExecuteMultipleResponse ExecuteMultiple(BulkAction bulkAction, IEnumerable<Entity> entityDataList)
        {
            try
            {
                var crmServiceClient = GetCrmServiceClient();
                if (crmServiceClient != null && crmServiceClient.IsReady)
                {
                    var request = new ExecuteMultipleRequest()
                    {
                        Requests = new OrganizationRequestCollection(),
                        Settings = new ExecuteMultipleSettings
                        {
                            ContinueOnError = false,
                            ReturnResponses = true
                        }
                    };

                    switch (bulkAction)
                    {
                        case BulkAction.Create:
                            {
                                foreach (var entityData in entityDataList)
                                {
                                    request.Requests.Add(new CreateRequest() { Target = entityData });
                                }
                                break;
                            }
                        case BulkAction.Update:
                            {
                                foreach (var entityData in entityDataList)
                                {
                                    request.Requests.Add(new UpdateRequest() { Target = entityData });
                                }
                                break;
                            }
                        default:
                            break;
                    }


                    var response = (ExecuteMultipleResponse)crmServiceClient.Execute(request);
                    return response;
                }
                return new ExecuteMultipleResponse();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

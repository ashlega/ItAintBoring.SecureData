using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace ITAintBoring.SecureData.Plugins
{
    public class SecureDataPlugin : IPlugin
    {
        public void ExecuteRetrieve(IOrganizationService service, IPluginExecutionContext context)
        {
            Entity entity = (Entity)context.OutputParameters["BusinessEntity"];
            try
            {
                Entity entWithSecData = service.Retrieve(entity.LogicalName, entity.Id, new ColumnSet("ita_securedata"));
                if (entWithSecData.Contains("ita_securedata"))
                {
                    var secureDataReference = (EntityReference)entWithSecData["ita_securedata"];
                    var secureData = service.Retrieve(secureDataReference.LogicalName, secureDataReference.Id, new ColumnSet(true));
                    if (secureData.Contains("ita_details")) entity["description"] = secureData["ita_details"];
                }
            }
            catch(Exception ex)
            {
                //security error - do nothing
                entity["description"] = "This is a protected email - please contact the owner to get access!";
            }
        }

       
        public void ExecuteUpdateCreate(IOrganizationService service, IPluginExecutionContext context)
        {
            string protectedEmail = "This is a protected email!";

            Entity target = (Entity)context.InputParameters["Target"];
            Entity postImage = null;
            if (context.PostEntityImages.Contains("PostImage")) postImage = (Entity)context.PostEntityImages["PostImage"];
            Entity preImage = null;
            if (context.PreEntityImages.Contains("PreImage")) preImage = (Entity)context.PreEntityImages["PreImage"];

            //Switching isSecure
            bool isSecureChanged = target.Contains("ita_issecure") &&
                                      (preImage == null && (bool)target["ita_issecure"] == true
                                       || preImage != null && preImage.Contains("ita_issecure") && (bool)preImage["ita_issecure"] == true && target["ita_issecure"] == null
                                       || preImage != null && target["ita_issecure"] != null && (bool)preImage["ita_issecure"] != (bool)target["ita_issecure"]);
            bool isSecure = target.Contains("ita_issecure") && target["ita_issecure"] != null && (bool)target["ita_issecure"] == true;
            Entity secureData = null;

            if (isSecureChanged)
            {
                if (isSecure)
                {
                    secureData = new Entity("ita_securedata");
                    secureData["ita_details"] = target.Contains("description") ? target["description"] : (preImage.Contains("description") ? preImage["description"] : null);
                    secureData.Id = service.Create(secureData);
                    target["ita_securedata"] = secureData.ToEntityReference();
                    target["description"] = protectedEmail;
                }
                else 
                {
                    EntityReference secureDataRef = (EntityReference)preImage["ita_securedata"];
                    secureData = service.Retrieve(secureDataRef.LogicalName, secureDataRef.Id, new ColumnSet(true));
                    service.Delete(secureData.LogicalName, secureData.Id);
                    if (!target.Contains("description")) //if there is an updated description, no need to restore from the secure record
                        target["description"] = secureData["ita_details"];
                }
            }
            else if(target.Contains("description"))
            {
                if (preImage.Contains("ita_securedata"))
                {
                    secureData = new Entity("ita_securedata");
                    secureData["ita_details"] = target["description"];
                    secureData.Id = ((EntityReference)preImage["ita_securedata"]).Id;
                    service.Update(secureData);
                }
            }
            


            //Changing statuses
            if (target.Contains("statuscode"))
            {
                int newStatus = ((OptionSetValue)target["statuscode"]).Value;
                if(newStatus == 6 || newStatus == 7)
                {
                    //Need to unhide the "description" so outlook integration works properly
                    if(preImage == null 
                        || preImage.Contains("ita_issecure") && (bool)preImage["ita_issecure"] == true && (!target.Contains("ita_issecure") || isSecure)
                        || !preImage.Contains("ita_issecure") && isSecure)
                    {
                        if(secureData == null)
                        {
                            EntityReference secureDataRef = (EntityReference)preImage["ita_securedata"];
                            secureData = service.Retrieve(secureDataRef.LogicalName, secureDataRef.Id, new ColumnSet(true));
                        }
                        target["description"] = secureData["ita_details"];
                    }
                }
                else
                {
                    //Need to hide the description again
                    if (preImage == null
                        || preImage.Contains("ita_issecure") && (bool)preImage["ita_issecure"] == true && (!target.Contains("ita_issecure") || isSecure)
                        || !preImage.Contains("ita_issecure") && isSecure)
                    {
                        target["description"] = protectedEmail;
                    }
                }
            }
        }
        

        public void Execute(IServiceProvider serviceProvider)
        {
            
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.Stage == 40)
            {
                if (context.MessageName == "Retrieve")
                {
                    ExecuteRetrieve(service, context);
                }
            }
            else if (context.MessageName == "Update" || context.MessageName == "Create")
            {
                ExecuteUpdateCreate(service, context);
            }

        }


    }
}

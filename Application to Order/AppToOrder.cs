using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace Application_to_Order
{
    public class AppToOrder : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));
            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is EntityReference)
            {
                // Obtain the target entity from the input parameters.
                EntityReference reference = (EntityReference)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                try
                {
                    Entity entity = new Entity();
                    entity = service.Retrieve(reference.LogicalName, reference.Id, new ColumnSet(true));// ("ss_application");


                    string fetchXml = @"<fetch distinct='false' mapping='logical' returntotalrecordcount='true' page='1' count='50' no-lock='false'>
                                     <entity name='salesorder'>
                                         <attribute name='entityimage_url'/>
                                         <attribute name='name'/>
                                         <attribute name='statecode'/>
                                         <attribute name='totalamount'/>
                                         <attribute name='salesorderid'/>
                                         <order attribute='name' descending='true'/>
                                         <link-entity name='ss_application' from='ss_applicationid' to='ss_application' alias='bb'>
                                             <filter type='and'>
                                                <condition attribute='ss_applicationid' operator='eq' uitype='ss_application' value='" + entity.Id + @"'/>
                                             </filter>  
                                         </link-entity>
                                     </entity>
                                 </fetch>";

                    EntityCollection appOrders = service.RetrieveMultiple(new FetchExpression(fetchXml));

                    if (appOrders.TotalRecordCount == 0)
                    {

                        Guid orderId = Guid.NewGuid();
                        Entity order = new Entity("salesorder", orderId);
                        Entity orderLine = new Entity("salesorderdetail");

                        if (entity.Contains("ss_applicationid"))
                        {
                            order["name"] = ("Order for " + entity.Attributes["ss_name"]);
                            order["ss_application"] = new EntityReference("ss_application", entity.Id);
                            order["pricelevelid"] = entity.Attributes["ss_pricelist"];
                            order["customerid"] = entity.Attributes["ss_customer"];
                            order["ss_destinationaddress"] = entity.Attributes["ss_destinationaddress"];

                            // Create the Order in Microsoft Dynamics CRM.
                            tracingService.Trace("AppOrderPlugin: Creating the Order.");
                            service.Create(order);

                            orderLine["isproductoverridden"] = false;
                            orderLine["productid"] = entity.Attributes["ss_product"];
                            orderLine["uomid"] = new EntityReference("uom", Guid.Parse("46d8b737-2339-4011-984a-5e54126ccdb2")); //uoms
                            orderLine["salesorderid"] = new EntityReference("salesorder", orderId);
                            orderLine["quantity"] = Convert.ToDecimal(1);

                            // Create the Order Line in Microsoft Dynamics CRM.
                            tracingService.Trace("AppOrderPlugin: Creating the Product Order Line.");
                            service.Create(orderLine);

                            if (entity.FormattedValues["ss_applicationtype"].Equals("Package Submission"))
                            {
                                Entity shippingLine = new Entity("salesorderdetail");
                                Entity primaryUnit = new Entity("uom");


                                ConditionExpression condition = new ConditionExpression
                                {
                                    AttributeName = "name",
                                    Operator = ConditionOperator.Equal
                                };
                                condition.Values.Add("Primary Unit");

                                FilterExpression filter = new FilterExpression();
                                filter.AddCondition(condition);

                                QueryExpression query = new QueryExpression("uom");
                                query.ColumnSet.AddColumn("uomid");
                                query.Criteria.AddFilter(filter);

                                EntityCollection uomList = service.RetrieveMultiple(query);

                                //@"<Fetch mapping='logical'>
                                //    <entity name='uom'>
                                //        <attribute name='uomid'/>
                                //        <attribute name='name'/>
                                //        <filter type='and'>   
                                //            <condition attribute='name' operator='eq' value='Primary Unit'/>   
                                //        </filter>   
                                //    </entity>
                                //</fetch>";

                                if (uomList.TotalRecordCount > 0)
                                {
                                    Guid uomId = uomList[0].Id;                      

                                    shippingLine["isproductoverridden"] = false;
                                    shippingLine["productid"] = entity.Attributes["ss_shippingspeed"];
                                    shippingLine["uomid"] = new EntityReference("uom", uomId); //Guid.Parse("46d8b737-2339-4011-984a-5e54126ccdb2") /uoms 
                                    shippingLine["salesorderid"] = new EntityReference("salesorder", orderId);
                                    shippingLine["quantity"] = Convert.ToDecimal(1); ;

                                    // Create the Order Line in Microsoft Dynamics CRM.
                                    tracingService.Trace("AppOrderPlugin: Creating the Shipping Speed Order Line.");
                                    service.Create(shippingLine);
                                }

                                // Close Application
                                if (order.Contains("salesorderid"))
                                {
                                    entity["statuscode"] = 1;
                                    entity["statecode"] = 2;
                                }
                            }
                        }
                    } else
                    {
                        throw new InvalidPluginExecutionException("An order already exists for this application");
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("FollowUpPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}

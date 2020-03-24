using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;

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
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];
                // Obtain the organization service reference which you will need for
                // web service calls.
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
                try
                {
                    Guid orderId = Guid.NewGuid();
                    Entity order = new Entity("salesorder", orderId);
                    Entity orderLine = new Entity("salesorderdetail");

                    if (context.OutputParameters.Contains("id"))
                    {
                        order["name"] = ("Order for " + entity.Attributes["ss_name"]);
                        order["ss_application"] = new EntityReference("ss_application", entity.Id);
                        order["pricelevelid"] = entity.Attributes["ss_pricelist"];
                        order["customerid"] = entity.Attributes["ss_customer"];

                        // Create the Order in Microsoft Dynamics CRM.
                        tracingService.Trace("FollowupPlugin: Creating the Order.");
                        service.Create(order);

                        orderLine["isproductoverridden"] = false;
                        orderLine["productid"] = entity.Attributes["ss_product"];
                        orderLine["uomid"] = new EntityReference("uom", Guid.Parse("46d8b737-2339-4011-984a-5e54126ccdb2")); //uoms
                        orderLine["salesorderid"] = new EntityReference("salesorder", orderId);
                        orderLine["quantity"] = Convert.ToDecimal(1);

                        // Create the Order Line in Microsoft Dynamics CRM.
                        tracingService.Trace("AppOrderPlugin: Creating the Order Line.");
                        service.Create(orderLine);

                        if (entity.FormattedValues["ss_applicationtype"].Equals("Package Submission"))
                        {
                            Entity shippingLine = new Entity("salesorderdetail");
                            shippingLine["isproductoverridden"] = false;
                            shippingLine["productid"] = entity.Attributes["ss_shippingspeed"];
                            shippingLine["uomid"] = new EntityReference("uom", Guid.Parse("46d8b737-2339-4011-984a-5e54126ccdb2")); //uoms
                            shippingLine["salesorderid"] = new EntityReference("salesorder", orderId);
                            shippingLine["quantity"] = Convert.ToDecimal(1); ;

                            // Create the Order Line in Microsoft Dynamics CRM.
                            tracingService.Trace("AppOrderPlugin: Creating the Order Line.");
                            service.Create(shippingLine);

                        }
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

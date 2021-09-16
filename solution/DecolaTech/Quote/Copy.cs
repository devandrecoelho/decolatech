using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Actions.Quote
{
    public class Copy : CodeActivity
    {
        [RequiredArgument]
        [Input("quoteId")]
        public InArgument<string> quoteId { get; set; }

        [Output("copyQuoteId")]
        public OutArgument<string> copyQuoteId { get; set; }

        readonly List<string> systemsKeyFieldsList = new List<string>();
        protected override void Execute(CodeActivityContext context)
        {

            #region Declarations
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService orgService = serviceFactory.CreateOrganizationService(workflowContext.UserId);
            ITracingService tracingService = (ITracingService)context.GetExtension<ITracingService>();
            #endregion

            tracingService.Trace("Starting Copy Quote For IdCotacao = " + quoteId.Get(context));

            try
            {
                #region QuoteFields
                ColumnSet quoteFields = new ColumnSet(
                    "name",
                    "transactioncurrencyid",
                    "pricelevelid",
                    "shippingmethodcode",
                    "paymenttermscode",
                    "freighttermscode",
                    "billto_line1",
                    "billto_line2",
                    "billto_line3",
                    "billto_city",
                    "billto_stateorprovince",
                    "billto_postalcode",
                    "billto_country",
                    "willcall",
                    "shipto_line1",
                    "shipto_line2",
                    "shipto_line3",
                    "shipto_city",
                    "shipto_stateorprovince",
                    "shipto_postalcode",
                    "shipto_country",
                    "discountpercentage",
                    "discountamount",
                    "freightamount",
                    "opportunityid",
                    "customerid",
                    "description"
                    );
                #endregion

                tracingService.Trace("Retrieving origin quote fields");
                Entity quoteOrigin = orgService.Retrieve("quote", new Guid(quoteId.Get(context)), quoteFields);
                Entity quoteCopy = new Entity("quote");

                tracingService.Trace("Setting uncopy fields");
                SetSystemKeyFieldsList();

                foreach (var attributeOrigin in quoteOrigin.Attributes)
                {

                    if (!systemsKeyFieldsList.Contains(attributeOrigin.Key))
                    {
                        quoteCopy.Attributes[attributeOrigin.Key] = attributeOrigin.Value;
                    }

                }

                tracingService.Trace("Create a Quote Copy");
                Guid newQuote = orgService.Create(quoteCopy);

                tracingService.Trace("Retrieving Origin Quote Details");
                EntityCollection quoteDetailsOrigin = RetrieveQuoteDetail(quoteOrigin.Id, orgService, tracingService);

                if (quoteDetailsOrigin.Entities.Count > 0)
                {
                    tracingService.Trace("Starting loop to copy quote details");
                    foreach (var quoteDetailOrigin in quoteDetailsOrigin.Entities)
                    {
                        Entity QuoteDetailCopy = new Entity("quotedetail");

                        tracingService.Trace("Setting quote id to a new Quote Detail");
                        QuoteDetailCopy.Attributes["quoteid"] = new EntityReference("quote", newQuote);

                        tracingService.Trace("Copying Attributes to a new Quote Detail");
                        foreach (var quoteDetalAttributeOrigin in quoteDetailOrigin.Attributes)
                        {
                            if (!systemsKeyFieldsList.Contains(quoteDetalAttributeOrigin.Key))
                            {
                                QuoteDetailCopy.Attributes[quoteDetalAttributeOrigin.Key] = quoteDetalAttributeOrigin.Value;
                            }
                        }

                        tracingService.Trace("Creating a new Quote detail");
                        Guid quoteDetailCopy = orgService.Create(QuoteDetailCopy);

                        tracingService.Trace($"Return Quote Detail id {quoteDetailCopy}");
                    }
                }

                tracingService.Trace($"Return Copy Quote id {newQuote}");
                copyQuoteId.Set(context, newQuote.ToString());

            }
            catch (Exception e)
            {
                throw new InvalidOperationException(e.Message);
            }

        }

        private EntityCollection RetrieveQuoteDetail(Guid quoteId, IOrganizationService orgService, ITracingService tracingService)
        {
            tracingService.Trace("Starting RetrieveQuoteDetail Method");

            #region QuoteDetailFields
            ColumnSet quoteDetailFields = new ColumnSet(
                "productid",
                "uomid",
                "ispriceoverridden",
                "quantity",
                "manualdiscountamount",
                "tax",
                "quoteid",
                "requestdeliveryby",
                "salesrepid",
                "willcall",
                "shipto_name",
                "shipto_line1",
                "shipto_line2",
                "shipto_line3",
                "shipto_city",
                "shipto_stateorprovince",
                "shipto_postalcode",
                "shipto_country",
                "shipto_telephone",
                "shipto_freighttermscode",
                "shipto_contactname"
                );
            #endregion

            QueryExpression queryQuoteDetail = new QueryExpression("quotedetail")
            {
                EntityName = "quotedetail",
                ColumnSet = quoteDetailFields,
                Criteria =
                {
                    Filters =
                    {
                        new FilterExpression()
                        {
                            Conditions =
                            {
                                new ConditionExpression("quoteid", ConditionOperator.Equal, quoteId)
                            }
                        }
                    }
                }
            };

            tracingService.Trace("Querying Quote Details");
            EntityCollection quoteDetailOrigin = orgService.RetrieveMultiple(queryQuoteDetail);

            tracingService.Trace($"Found {quoteDetailOrigin.Entities.Count} Records.");
            return quoteDetailOrigin;
        }
        private void SetSystemKeyFieldsList()
        {
            systemsKeyFieldsList.Add("quoteid");
            systemsKeyFieldsList.Add("quotedetailid");
        }
    }
}

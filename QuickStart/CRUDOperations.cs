using System;
using System.ServiceModel;
using System.IdentityModel.Tokens;
using System.ServiceModel.Security;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System.Diagnostics;
using System.Collections.Generic;

namespace Microsoft.Crm.Sdk.Samples
{
    public class CRUDOperations
    {
        static OrganizationServiceProxy serviceProxy = null;
        static IOrganizationService service = null;

        static string _customEntityName = "new_customentity";
        static string _customActivityName = "new_customactivity";

        static public void Main(string[] args)
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();

                //Change below 3 lines as per your env. 
                Uri discoveryServerUri = new Uri("http://onefarm7508.onefarm7508dom.extest.microsoft.com/CITTest/XRMServices/2011/Organization.svc");
                credentials.UserName.UserName = "~\administrator";
                credentials.UserName.Password = "T!T@n1130";

                serviceProxy = new OrganizationServiceProxy(discoveryServerUri, null, credentials, null);
				serviceProxy.EnableProxyTypes(); 
                service = (IOrganizationService)serviceProxy;

                //Create an Account with multiple Contacts, so that it shows in associated grid. 
                //CreateOneAccountWithMultipleContacts(1, 70);

                //********** FOR ACCOUNT **********
                //To create Account. One can implement method similar to CreateAccounts for other entities like CreateContacts. 
                CreateAccounts(1, 5, 10);

                //If needed to delete above created records (Can also delete some of them), uncomment below line. 
                //DeleteAccountRecords(1, 5);


                //********** FOR OPPORTUNITY ********** 
                //To create opportunity 
               // CreateOpportunities(1, 70);

                //If needed to delete, then uncomment below lines. 
               // DeleteOpportunityRecords(1, 5);


                //********** FOR CUSTOM ENTITY ********** 
                //If you need to create Custom Entity uncomment below lines. Next time comment creation of customEntity. 
               // CreateCustomEntity();

                //Create customEntity records 
                //CreateCustomEntityRecords(1, 70);

                //To delete customEntity records created using this tool, uncomment below line
               // DeleteCustomEntityRecords(1, 5);

                //If you want to delete the CustomEntity itself, which you created earlier using this tool, then uncomment below line
               // DeleteCustomEntity();

                //********** FOR CUSTOM ACTIVITY ********** 
                //If you need to create Custom Activity records uncomment below lines. Next time comment creation of customActivity. 
               // CreateCustomActivity();

                //Create customActivity records 
                //CreateCustomActivityRecords(1, 70);

                //To delete customActivity records created using this tool, uncomment below line
                //DeleteCustomActivityRecords(1, 70);

                //If you want to delete the CustomActivity itself, which you created earlier using this tool 
                //DeleteCustomActivity();

            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> e) { HandleException(e); }
            catch (TimeoutException e) { HandleException(e); }
            catch (SecurityTokenValidationException e) { HandleException(e); }
            catch (ExpiredSecurityTokenException e) { HandleException(e); }
            catch (MessageSecurityException e) { HandleException(e); }
            catch (SecurityNegotiationException e) { HandleException(e); }
            catch (SecurityAccessDeniedException e) { HandleException(e); }
            catch (Exception e) { HandleException(e); }
            finally
            {
                // Always dispose the service object to close the service connection and free resources.
                if (serviceProxy != null) serviceProxy.Dispose();

                Console.WriteLine("Press <Enter> to exit.");
                Console.ReadLine();
            }
        }

        /// Handle a thrown exception.
        /// </summary>
        /// <param name="ex">An exception.</param>
        public static void HandleException(Exception e)
        {
            // Display the details of the exception.
            Console.WriteLine("\n" + e.Message);
            Console.WriteLine(e.StackTrace);

            if (e.InnerException != null) HandleException(e.InnerException);
        }

        #region Account Create and Delete
        /// <summary>
        /// Create Accounts having index appended to TestAccount from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void CreateAccounts(int startNumber, int endNumber, int sizeOfFieldValues)
        {
			var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
			var stringChars = new char[sizeOfFieldValues];
			var random = new Random();

			for (int i = 0; i < stringChars.Length; i++)
			{
				stringChars[i] = chars[random.Next(chars.Length)];
			}

			var finalString = new String(stringChars);
			Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Accounts");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Account account = new Account();
                account.Name = "TestAccount" + finalString + i;
                account.Telephone1 = "123587258787875738275718725385381735753821758324432432" + i;
                //Currency 
                account.Revenue = new Money(new decimal(2000 + i));
				account.AccountCategoryCode = new OptionSetValue((int)AccountAccountCategoryCode.PreferredCustomer);
				
				account.EMailAddress1 = finalString + i + "@xyz.com";
				account.Fax = finalString + i;
				account.NumberOfEmployees = 100 + i;
				account.WebSiteURL = "www." + finalString + ".com"; 
                //Datetime 
                account.LastUsedInCampaign = DateTime.Now.AddMonths(-i);
                //Floating point 
                account.Address1_Latitude = (15.38 + i) % 90;

                Lead lead = new Lead();
                lead.LastName = "TestOriginatingLeadForAccountLN" + finalString+ i;
                lead.FirstName = "TestFNO" + i;
                Guid leadId = service.Create(lead);

                //Lookup 
                account.OriginatingLeadId = new EntityReference(Lead.EntityLogicalName, leadId);
                //MultiLine
                account.Description = "This is description with \n newline \n again \n" + finalString + "Some other random things" + i;
				//OptionSet
				account.CustomerTypeCode = new OptionSetValue((int)AccountCustomerTypeCode.Investor);
                //Two option
                account.DoNotEMail = false;

                //Whole number 
                account.SharesOutstanding = i;
				account.Telephone3 = finalString + i; 
                Guid accountId = service.Create(account);
                Console.WriteLine("Created Account: TestAccount" + i + " having guid: " + accountId);
            }
        }

        /// <summary>
        /// Delete Account records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void DeleteAccountRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Deleting Account records between " + startNumber + " and " + endNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                //Retrieve guid of TestAccount1 - say Entity is Account and startNumber is 1. 
                ColumnSet cols = new ColumnSet(true);
                Guid? accountId = GetGuiByAttributeValue("name", "TestAccount" + i, Account.EntityLogicalName);
                if (accountId == null)
                {
                    Console.WriteLine("Error : Did not return Guid for TestAccount" + i + ", May be its deleted already?");
                    continue;
                }
                Account retrievedAccount = (Account)service.Retrieve("account", (Guid)accountId, cols);
                Console.WriteLine("Deleting TestAccount" + i);
                service.Delete(Account.EntityLogicalName, (Guid)accountId);
            }
        }

        #endregion

        #region Opportunity Create and Delete
        /// <summary>
        /// Create Opportunities having index appended to TestOpportunity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void CreateOpportunities(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Opportunities");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Opportunity opportunity = new Opportunity();
                opportunity.Name = "TestOpportunity" + i;
                opportunity.EstimatedValue = new Money(new Decimal(500.23));
                Guid opportunityId = service.Create(opportunity);
                Console.WriteLine("Created Opportunity: TestOpportunity" + i + " having guid: " + opportunityId);
            }
        }

        /// <summary>
        /// Delete Opportunity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void DeleteOpportunityRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Deleting Opportunity records between " + startNumber + " and " + endNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                //Retrieve guid of TestOpportunity1 - say Entity is Opportunity and startNumber is 1. 
                ColumnSet cols = new ColumnSet(true);
                Guid? opportunityId = GetGuiByAttributeValue("name", "TestOpportunity" + i, Opportunity.EntityLogicalName);
                if (opportunityId == null)
                {
                    Console.WriteLine("Error : Did not return Guid for TestOpportunity" + i + ", May be its deleted already?");
                    continue;
                }
                Opportunity retrievedOpportunity = (Opportunity)service.Retrieve("opportunity", (Guid)opportunityId, cols);
                Console.WriteLine("Deleting TestOpportunity" + i);
                service.Delete(Opportunity.EntityLogicalName, (Guid)opportunityId);
            }
        }

        #endregion

        #region Custom Entity Create and Delete
        /// <summary>
        /// Create Custom Entity having index appended to TestCustomEntity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void CreateCustomEntityRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Custom Entities");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Entity customEntity = new Entity(_customEntityName);
                customEntity["new_name"] = "TestCustomEntity" + i;
                customEntity["new_teststring"] = "Test String Data " + i;
                customEntity["new_testmoney"] = new Money(new decimal(532.23));
                customEntity["new_testdate"] = DateTime.Now;

                //Create a lead for reference 
                Lead lead = new Lead();
                lead.Subject = "TestLeadRefCustom" + i;
                lead.FirstName = "TestLeadRefFirstName" + i;
                lead.LastName = "TestLeadRefLastName" + i;
                Guid leadId = service.Create(lead);
                customEntity["new_customentity_leadid"] = new EntityReference(_customEntityName, leadId);

                Guid customEntityId = service.Create(customEntity);
                Console.WriteLine("Created Custom Entity: TestCustomEntity" + i + " having guid: " + customEntityId);
            }
        }

        /// <summary>
        /// Delete Custom Entity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void DeleteCustomEntityRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Deleting Custom Entity records between " + startNumber + " and " + endNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                //Retrieve guid of TestCustomEntity1 - say Entity is CustomEntity and startNumber is 1. 
                ColumnSet cols = new ColumnSet(true);
                Guid? customEntityId = GetGuiByAttributeValue("new_name", "TestCustomEntity" + i, _customEntityName);
                if (customEntityId == null)
                {
                    Console.WriteLine("Error : Did not return Guid for TestCustomEntity" + i + ", May be its deleted already?");
                    continue;
                }
                Entity retrievedCustomEntity = service.Retrieve(_customEntityName, (Guid)customEntityId, cols);
                Console.WriteLine("Deleting TestCustomEntity" + i);
                service.Delete(_customEntityName, (Guid)customEntityId);
            }
        }

        #endregion

        #region Custom Activity Create and Delete
        /// <summary>
        /// Create Custom Activities having index appended to TestCustomActivity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void CreateCustomActivityRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Custom Activities");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Entity customActivity = new Entity(_customActivityName);
                customActivity["subject"] = "TestCustomActivity" + i;
                customActivity["new_teststring"] = "Test String Activity Data " + i;
                customActivity["new_testinteger"] = 500 + i;

                //Create a contact for regarding object id for activity
                Contact contact = new Contact();
                contact.LastName = "TestContactRefLastName" + i;
                contact.FirstName = "TestContactRefFirstName" + i;
                Guid contactId = service.Create(contact);
                customActivity["regardingobjectid"] = new EntityReference(Contact.EntityLogicalName, contactId);

                Guid customActivityId = service.Create(customActivity);
                Console.WriteLine("Created Custom Activity: TestCustomActivity" + i + " having guid: " + customActivityId);
            }
        }


        /// <summary>
        /// Delete Custom Activity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public static void DeleteCustomActivityRecords(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Deleting Custom Activity records between " + startNumber + " and " + endNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                //Retrieve guid of TestCustomActvity1 - say Entity is CustomActivity and startNumber is 1. 
                ColumnSet cols = new ColumnSet(true);

                Guid? customActivityId = GetGuiByAttributeValue("subject", "TestCustomActivity" + i, _customActivityName);
                if (customActivityId == null)
                {
                    Console.WriteLine("Error : Did not return Guid for TestCustomActivity" + i + ", May be its deleted already?");
                    continue;
                }
                Entity retrievedCustomActivity = service.Retrieve(_customActivityName, (Guid)customActivityId, cols);
                Console.WriteLine("Deleting TestCustomActivity" + i);
                service.Delete(_customActivityName, (Guid)customActivityId);
            }
        }

        #endregion


        /// <summary>
        /// Common method to be used to get the Guid of record, when guid is not available, but some other attribute value is available. 
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="attributeValue"></param>
        /// <param name="entityLogicalName"></param>
        /// <returns></returns>
        public static Guid? GetGuiByAttributeValue(string attributeName, string attributeValue, string entityLogicalName)
        {
            QueryExpression query = new QueryExpression(entityLogicalName);
            FilterExpression flterExp = new FilterExpression();
            flterExp.Conditions.Add(new ConditionExpression(attributeName, ConditionOperator.Equal, attributeValue));
            query.Criteria = flterExp;
            if (entityLogicalName == _customActivityName)
            {
                query.ColumnSet = new ColumnSet("activityid");
            }
            else
            {
                query.ColumnSet = new ColumnSet(entityLogicalName + "id");
            }

            EntityCollection entityCollection = null;
            entityCollection = service.RetrieveMultiple(query);
            if (entityCollection.Entities.Count >= 1)
            {
                return entityCollection.Entities[0].Id;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Create a custom entity
        /// </summary>
        /// <returns></returns>
        public static void CreateCustomEntity()
        {
            Console.WriteLine("###### Creating Custom Entity");
            CreateEntityRequest createrequest = new CreateEntityRequest
            {

                //Define the entity. This entity will show up in Sales/Service/Marketing -> Extensions (last group) 
                Entity = new EntityMetadata
                {
                    SchemaName = _customEntityName,
                    DisplayName = new Label("Custom Entity", 1033),
                    DisplayCollectionName = new Label("Custom Entities", 1033),
                    Description = new Label("A custom entity created for testing purpose", 1033),
                    OwnershipType = OwnershipTypes.UserOwned,
                    IsActivity = false,

                },

                // Define the primary attribute for the entity
                PrimaryAttribute = new StringAttributeMetadata
                {
                    SchemaName = "new_name",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    FormatName = StringFormatName.Text,
                    DisplayName = new Label("Name", 1033),
                    Description = new Label("The primary attribute for the Custom Entity.", 1033)
                }

            };
            service.Execute(createrequest);

            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);

            //Add String type attribute 
            CreateAttributeRequest createBankNameAttributeRequest = new CreateAttributeRequest
            {
                EntityName = _customEntityName,
                Attribute = new StringAttributeMetadata
                {
                    SchemaName = "new_teststring",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    FormatName = StringFormatName.Text,
                    DisplayName = new Label("Test String", 1033)
                }
            };
            service.Execute(createBankNameAttributeRequest);

            //Add money type attribute 
            CreateAttributeRequest createBalanceAttributeRequest = new CreateAttributeRequest
            {
                EntityName = _customEntityName,
                Attribute = new MoneyAttributeMetadata
                {
                    SchemaName = "new_testmoney",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    PrecisionSource = 2,
                    DisplayName = new Label("Test Money", 1033)
                }
            };
            service.Execute(createBalanceAttributeRequest);

            //Add date only attribute 
            CreateAttributeRequest createCheckedDateRequest = new CreateAttributeRequest
            {
                EntityName = _customEntityName,
                Attribute = new DateTimeAttributeMetadata
                {
                    SchemaName = "new_testdate",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    Format = DateTimeFormat.DateOnly,
                    DisplayName = new Label("Test Date", 1033),
                }
            };
            service.Execute(createCheckedDateRequest);

            //Add lookup attribute 
            CreateOneToManyRequest req = new CreateOneToManyRequest()
            {
                Lookup = new LookupAttributeMetadata()
                {
                    Description = new Label("The referral (lead) from the Custom Entity record", 1033),
                    DisplayName = new Label("Test Lookup To Lead", 1033),
                    LogicalName = "new_customentity_leadid",
                    SchemaName = "new_customentity_leadid",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.Recommended)
                },
                OneToManyRelationship = new OneToManyRelationshipMetadata()
                {
                    AssociatedMenuConfiguration = new AssociatedMenuConfiguration()
                    {
                        Behavior = AssociatedMenuBehavior.UseCollectionName,
                        Group = AssociatedMenuGroup.Details,
                        Label = new Label("Custom Entities", 1033),
                        Order = 10000
                    },
                    CascadeConfiguration = new CascadeConfiguration()
                    {
                        Assign = CascadeType.Cascade,
                        Delete = CascadeType.Cascade,
                        Merge = CascadeType.Cascade,
                        Reparent = CascadeType.Cascade,
                        Share = CascadeType.Cascade,
                        Unshare = CascadeType.Cascade
                    },
                    ReferencedEntity = "lead",
                    ReferencedAttribute = "leadid",
                    ReferencingEntity = _customEntityName,
                    SchemaName = "new_customentity_leadid"
                }
            };
            service.Execute(req);

            // Customizations must be published after an entity is updated.
            publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);
            Console.WriteLine("Created Custom Entity");
        }

        /// <summary>
        /// Delete the Custom Entity itself
        /// </summary>
        public static void DeleteCustomEntity()
        {
            DeleteEntityRequest request = new DeleteEntityRequest()
            {
                LogicalName = _customEntityName,
            };
            service.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);
            Console.WriteLine("The custom entity has been deleted.");
        }

        /// <summary>
        /// Create a custom activity
        /// </summary>
        /// <returns></returns>
        public static void CreateCustomActivity()
        {
            Console.WriteLine("###### Creating Custom Activity");
            // Create the custom activity entity.
            CreateEntityRequest request = new CreateEntityRequest
            {
                HasNotes = true,
                HasActivities = false,
                PrimaryAttribute = new StringAttributeMetadata
                {
                    SchemaName = "Subject",
                    RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                    MaxLength = 100,
                    DisplayName = new Label("Subject", 1033)
                },
                Entity = new EntityMetadata
                {
                    IsActivity = true,
                    SchemaName = _customActivityName,
                    DisplayName = new Label("Custom Activity", 1033),
                    DisplayCollectionName = new Label("Custom Activities", 1033),
                    OwnershipType = OwnershipTypes.UserOwned,
                    IsAvailableOffline = true,

                }
            };

            service.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);

            // Add few attributes to the custom activity entity.
            //Add string type attribute 
            CreateAttributeRequest stringAttributeRequest =
                new CreateAttributeRequest
                {
                    EntityName = _customActivityName,
                    Attribute = new StringAttributeMetadata
                    {
                        SchemaName = "new_teststring",
                        DisplayName = new Label("Test String", 1033),
                        MaxLength = 50
                    }
                };
            CreateAttributeResponse fontColorAttributeResponse =
                (CreateAttributeResponse)service.Execute(
                stringAttributeRequest);

            CreateAttributeRequest integerAttributeRequest =
                new CreateAttributeRequest
                {
                    EntityName = _customActivityName,
                    Attribute = new IntegerAttributeMetadata
                    {
                        SchemaName = "new_testinteger",
                        DisplayName = new Label("Test Integer", 1033)
                    }
                };
            CreateAttributeResponse fontSizeAttributeResponse =
                (CreateAttributeResponse)service.Execute(
                integerAttributeRequest);


            // Customizations must be published after an entity is updated.
            publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);
            Console.WriteLine("Created Custom Activity");
        }

        /// <summary>
        /// Delete the Custom Activity itself
        /// </summary>
        public static void DeleteCustomActivity()
        {
            DeleteEntityRequest request = new DeleteEntityRequest()
            {
                LogicalName = _customActivityName,
            };
            service.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            service.Execute(publishRequest);
            Console.WriteLine("The custom entity has been deleted.");
        }

        public static void CreateOneAccountWithMultipleContacts(int startNumber, int endNumber)
        {
            Account account = new Account();
            account.Name = "TestAccountWithMultipleContacts" + startNumber;
            Guid accountId = service.Create(account);
            List<Contact> listContacts = new List<Contact>();
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Contacts each having Account as TestAccountWithMultipleContacts" + startNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                Contact contact = new Contact();
                contact.LastName = "TestMultiContactLN" + i;
                contact.FirstName = "TestMultiContactFN" + i;
                Guid contactId = service.Create(contact);
                Contact retrievedContact = (Contact)service.Retrieve(Contact.EntityLogicalName, contactId, new ColumnSet(true));
                listContacts.Add(retrievedContact);
                Console.WriteLine("Created Contact: TestMultiContactLN" + i + " having guid: " + contactId);
            }
            Account retrievedAccount = (Account)service.Retrieve("account", accountId, new ColumnSet(true));
            retrievedAccount.contact_customer_accounts = listContacts;
            service.Update(retrievedAccount);
            Console.WriteLine("Added Account with " + (endNumber + 1 - startNumber) + " Contacts");
        }

    }
}

using Microsoft.Crm.QA.Utf;
using Microsoft.Crm.QA.Utils;
using System;
using Microsoft.Crm.Sdk.Samples;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Helper; 
namespace CommonScenarioTests
{
    [Microsoft.VisualStudio.TestTools.UnitTesting.TestClass]
    public sealed class Sales : WhoAmISdkTest
    {
        string _customEntityName = "new_customentity";
        string _customActivityName = "new_customactivity";
        /// <summary>
        /// Setup
        /// </summary>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestInitialize]
        public override void SetUp()
        {
            base.SetUp();
        }

        /// <summary>
        /// Treat this method as like Main() and perform all operations. Make sure to have corresponding methods. For some method calls Just Contact didn't work, so I had to use CrmSdk.Contact (For retrieve calls)
        /// Also 'static' from method copied from QuickStart Project was removed when used in this file. Replace service.  with Context.Proxy. in current document. 
        /// For this method to work, Make CreateDataInOnlineEnv as Startup Project and Change "CHANGE ONLY THIS SECTION" in App.config in this project. 
        /// </summary>
        [Microsoft.VisualStudio.TestTools.UnitTesting.TestMethod]
        public void VerifyNewAccountContactCreate()
        {
			CreateContacts(5328, 25000); 
			//CreationLib lib = new CreationLib();
			//lib.CreateAccount(this.Context.Proxy); 
			//Create an Account with multiple Contacts, so that it shows in associated grid. 
			//CreateOneAccountWithMultipleContacts(1, 70);

   //         //********** FOR ACCOUNT **********
   //         //To create Account. One can implement method similar to CreateAccounts for other entities like CreateContacts. 
   //         CreateAccounts(1, 70);

   //         //If needed to delete above created records (Can also delete some of them), uncomment below line. 
   //         DeleteAccountRecords(1, 5);

   //         //********** FOR OPPORTUNITY ********** 
   //         //To create opportunity 
   //         CreateOpportunities(1, 70);

   //         //If needed to delete, then uncomment below lines. 
   //         DeleteOpportunityRecords(1, 5);

   //         //********** FOR CUSTOM ENTITY ********** 
   //         //If you need to create Custom Entity uncomment below lines. Next time comment creation of customEntity. 
   //         CreateCustomEntity();

   //         //Create customEntity records 
   //         CreateCustomEntityRecords(1, 70);

   //         //To delete customEntity records created using this tool, uncomment below line
   //         DeleteCustomEntityRecords(1, 5);

   //         //If you want to delete the CustomEntity itself, which you created earlier using this tool, then uncomment below line
   //         //DeleteCustomEntity();

   //         //********** FOR CUSTOM ACTIVITY ********** 
   //         //If you need to create Custom Activity records uncomment below lines. Next time comment creation of customActivity. 
   //         CreateCustomActivity();

   //         //Create customActivity records 
   //         CreateCustomActivityRecords(1, 70);

   //         //To delete customActivity records created using this tool, uncomment below line
   //         DeleteCustomActivityRecords(1, 5);

            //If you want to delete the CustomActivity itself, which you created earlier using this tool 
            //DeleteCustomActivity();
        }

		/// <summary>
		/// Create Accounts having index appended to TestAccount from startNumber to endNumber
		/// </summary>
		/// <param name="startNumber"></param>
		/// <param name="endNumber"></param>
		public void CreateContacts(int startNumber, int endNumber)
		{
			Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Accounts");
			for (int i = startNumber; i <= endNumber; i++)
			{
				Contact contact = new Contact();
				contact.FirstName = "Test" + i;
				contact.LastName = "Contact" + i;
				contact.Salutation = "hi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hello";
				contact.JobTitle = "hi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hellohi, hello";
				contact.MobilePhone = "13211321132113211321132113211321132113211321";
				contact.Fax = "13211321132113211321132113211321132113211321";
				contact.EMailAddress1 = "hasgdkhasikal@ahfdilsudoadasdasfafsadfa.com";
				contact.Description = "Lorem ipsum is a pseudo-Latin text used in web design, typography, layout, and printing in place of English to emphasise design elements over content. It's also called placeholder (or filler) text. It's a convenient tool for mock-ups. It helps to outline the visual elements of a document or presentation, eg typography, font, or layout. Lorem ipsum is mostly a part of a Latin text by the classical author and philosopher Cicero. Its words and letters have been changed by addition or removal, so to deliberately render its content nonsensical; it's not genuine, correct, or comprehensible Latin anymore. While lorem ipsum's still resembles classical Latin, it actually has no meaning whatsoever. As Cicero's text doesn't contain the letters K, W, or Z, alien to latin, these, and others are often inserted randomly to mimic the typographic appearence of European languages, as are digraphs not to be found in the original. In a professional context it often happens that private or corporate clients corder a publication to be made and presented with the actual content still not being ready.Think of a news blog that's filled with content hourly on the day of going live. However, reviewers tend to be distracted by comprehensible content, say, a random text copied from a newspaper or the internet. The are likely to focus on the text, disregarding the layout and its elements. Besides, random text risks to be unintendedly humorous or offensive, an unacceptable risk in corporate environments. Lorem ipsum and its many variants have been employed since the early 1960ies, and quite likely since the sixteenth century. ";

				Guid contactId = this.Context.Proxy.Create(contact);
				Console.WriteLine("Created Contact: TestContact" + i + " having guid: " + contactId);
			}
		}

		public void CreateOneAccountWithMultipleContacts(int startNumber, int endNumber)
        {
            Account account = new Account();
            account.Name = "TestAccountWithMultipleContacts" + startNumber;
            Guid accountId = this.Context.Proxy.Create(account);
            CrmSdk.Account retrievedAccount = (CrmSdk.Account)this.Context.Proxy.Retrieve("account", accountId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
            List<CrmSdk.Contact> listContacts = new List<CrmSdk.Contact>();
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Contacts each having Account as TestAccountWithMultipleContacts" + startNumber);
            for (int i = startNumber; i <= endNumber; i++)
            {
                Contact contact = new Contact();
                contact.LastName = "TestMultiContactLN" + i;
                contact.FirstName = "TestMultiContactFN" + i;
                Guid contactId = this.Context.Proxy.Create(contact);
                CrmSdk.Contact retrievedContact = (CrmSdk.Contact)this.Context.Proxy.Retrieve(CrmSdk.Contact.EntityLogicalName, contactId, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
                listContacts.Add(retrievedContact);
                Console.WriteLine("Created Contact: TestMultiContactLN" + i + " having guid: " + contactId);
            }

            retrievedAccount.contact_customer_accounts = listContacts;
            this.Context.Proxy.Update(retrievedAccount);
            Console.WriteLine("Added Account with " + (endNumber + 1 - startNumber) + " Contacts");
        }

        /// <summary>
        /// Create Accounts having index appended to TestAccount from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void CreateAccounts(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Accounts");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Account account = new Account();
                account.Name = "TestAccount" + i;
                account.Telephone1 = "123" + i;
                //Currency 
                account.Revenue = new Money(new decimal(2000 + i));
                //Datetime 
                account.LastUsedInCampaign = DateTime.Now.AddMonths(-i);
                //Floating point 
                account.Address1_Latitude = (15.38 + i) % 90;

                Lead lead = new Lead();
                lead.LastName = "TestOriginatingLeadForAccountLN" + i;
                lead.FirstName = "TestFNO" + i;
                Guid leadId = Context.Proxy.Create(lead);

                //Lookup 
                account.OriginatingLeadId = new EntityReference(Lead.EntityLogicalName, leadId);
                //MultiLine
                account.Description = "Test description \n MultiLine 1 \n 2\n \n 3";
                //OptionSet
                account.CustomerTypeCode = new OptionSetValue((int)AccountCustomerTypeCode.Investor);
                //Two option
                account.DoNotEMail = false;

                //Whole number 
                account.SharesOutstanding = i;

                Guid accountId = Context.Proxy.Create(account);
                Console.WriteLine("Created Account: TestAccount" + i + " having guid: " + accountId);
            }
        }


        /// <summary>
        /// Create Opportunities having index appended to TestOpportunity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void CreateOpportunities(int startNumber, int endNumber)
        {
            Console.WriteLine("\n############# Creating " + (endNumber + 1 - startNumber) + " Opportunities");
            for (int i = startNumber; i <= endNumber; i++)
            {
                Opportunity opportunity = new Opportunity();
                opportunity.Name = "TestOpportunity" + i;
                opportunity.EstimatedValue = new Money(new Decimal(500.23));
                Guid opportunityId = Context.Proxy.Create(opportunity);
                Console.WriteLine("Created Opportunity: TestOpportunity" + i + " having guid: " + opportunityId);
            }
        }

        /// <summary>
        /// Create Custom Entity having index appended to TestCustomEntity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void CreateCustomEntityRecords(int startNumber, int endNumber)
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
                Guid leadId = Context.Proxy.Create(lead);
                customEntity["new_customentity_leadid"] = new EntityReference(_customEntityName, leadId);

                Guid customEntityId = Context.Proxy.Create(customEntity);
                Console.WriteLine("Created Custom Entity: TestCustomEntity" + i + " having guid: " + customEntityId);
            }
        }

        /// <summary>
        /// Create Custom Activities having index appended to TestCustomActivity from startNumber to endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void CreateCustomActivityRecords(int startNumber, int endNumber)
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
                Guid contactId = Context.Proxy.Create(contact);
                customActivity["regardingobjectid"] = new EntityReference(Contact.EntityLogicalName, contactId);

                Guid customActivityId = Context.Proxy.Create(customActivity);
                Console.WriteLine("Created Custom Activity: TestCustomActivity" + i + " having guid: " + customActivityId);
            }
        }
        /// <summary>
        /// Create a custom entity
        /// </summary>
        /// <returns></returns>
        public void CreateCustomEntity()
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
            Context.Proxy.Execute(createrequest);

            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);

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
            Context.Proxy.Execute(createBankNameAttributeRequest);

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
            Context.Proxy.Execute(createBalanceAttributeRequest);

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
            Context.Proxy.Execute(createCheckedDateRequest);

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
            Context.Proxy.Execute(req);

            // Customizations must be published after an entity is updated.
            publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);
            Console.WriteLine("Created Custom Entity");
        }

        /// <summary>
        /// Delete the Custom Entity itself
        /// </summary>
        public void DeleteCustomEntity()
        {
            DeleteEntityRequest request = new DeleteEntityRequest()
            {
                LogicalName = _customEntityName,
            };
            Context.Proxy.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);
            Console.WriteLine("The custom entity has been deleted.");
        }

        /// <summary>
        /// Create a custom activity
        /// </summary>
        /// <returns></returns>
        public void CreateCustomActivity()
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

            Context.Proxy.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);

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
                (CreateAttributeResponse)Context.Proxy.Execute(
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
                (CreateAttributeResponse)Context.Proxy.Execute(
                integerAttributeRequest);


            // Customizations must be published after an entity is updated.
            publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);
            Console.WriteLine("Created Custom Activity");
        }

        /// <summary>
        /// Delete the Custom Activity itself
        /// </summary>
        public void DeleteCustomActivity()
        {
            DeleteEntityRequest request = new DeleteEntityRequest()
            {
                LogicalName = _customActivityName,
            };
            Context.Proxy.Execute(request);
            // Customizations must be published after an entity is updated.
            PublishAllXmlRequest publishRequest = new PublishAllXmlRequest();
            Context.Proxy.Execute(publishRequest);
            Console.WriteLine("The custom entity has been deleted.");
        }

        /// <summary>
        /// Delete Account records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void DeleteAccountRecords(int startNumber, int endNumber)
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
                CrmSdk.Account retrievedAccount = (CrmSdk.Account)Context.Proxy.Retrieve("account", (Guid)accountId, cols);
                Console.WriteLine("Deleting TestAccount" + i);
                Context.Proxy.Delete(Account.EntityLogicalName, (Guid)accountId);
            }
        }

        /// <summary>
        /// Delete Opportunity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void DeleteOpportunityRecords(int startNumber, int endNumber)
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
                CrmSdk.Opportunity retrievedOpportunity = (CrmSdk.Opportunity)Context.Proxy.Retrieve("opportunity", (Guid)opportunityId, cols);
                Console.WriteLine("Deleting TestOpportunity" + i);
                Context.Proxy.Delete(Opportunity.EntityLogicalName, (Guid)opportunityId);
            }
        }

        /// <summary>
        /// Delete Custom Entity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void DeleteCustomEntityRecords(int startNumber, int endNumber)
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
                Entity retrievedCustomEntity = Context.Proxy.Retrieve(_customEntityName, (Guid)customEntityId, cols);
                Console.WriteLine("Deleting TestCustomEntity" + i);
                Context.Proxy.Delete(_customEntityName, (Guid)customEntityId);
            }
        }


        /// <summary>
        /// Delete Custom Activity records between startNumber and endNumber
        /// </summary>
        /// <param name="startNumber"></param>
        /// <param name="endNumber"></param>
        public void DeleteCustomActivityRecords(int startNumber, int endNumber)
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
                Entity retrievedCustomActivity = Context.Proxy.Retrieve(_customActivityName, (Guid)customActivityId, cols);
                Console.WriteLine("Deleting TestCustomActivity" + i);
                Context.Proxy.Delete(_customActivityName, (Guid)customActivityId);
            }
        }

        /// <summary>
        /// Common method to be used to get the Guid of record, when guid is not available, but some other attribute value is available. 
        /// </summary>
        /// <param name="attributeName"></param>
        /// <param name="attributeValue"></param>
        /// <param name="entityLogicalName"></param>
        /// <returns></returns>
        public Guid? GetGuiByAttributeValue(string attributeName, string attributeValue, string entityLogicalName)
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
            entityCollection = Context.Proxy.RetrieveMultiple(query);
            if (entityCollection.Entities.Count >= 1)
            {
                return entityCollection.Entities[0].Id;
            }
            else
            {
                return null;
            }
        }

    }
}

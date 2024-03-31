using Microsoft.Crm.Sdk.Messages;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IdentityModel.Metadata;
using System.Web.Services.Description;
using OfficeOpenXml;
using System.IO;

namespace UtilDuplication
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                DuplicateAcInvent();
                Console.WriteLine(" ========== De-Duplication Utility ============ ");
                Console.WriteLine("\n");
                Console.WriteLine(" ---------------------------------------------- ");
                Console.WriteLine("\n");

                Console.WriteLine("Connecting with A2C Production CRM");
                Console.WriteLine("\n");
                string connectionString = GetConnectionString();
                using (CrmServiceClient conn = new CrmServiceClient(connectionString))
                {
                    Console.WriteLine("Connection successful");
                    Console.WriteLine("\n");
                    Console.WriteLine("Fetching Results");
                    Console.WriteLine("\n");
                    int count = 0;

                    // Cast the proxy client to the IOrganizationService interface.
                    IOrganizationService orgService = (IOrganizationService)conn.OrganizationWebProxyClient ?? conn.OrganizationServiceProxy;

                    string fetchXml = GetFetchXml();
                    string fetchXmlforAccountID = GetFetchXmlForAccountID();
                    string fetchXmlforKYCRequired = GetFetchXmlForKYCRequired();
                    string fetchXmlforInventSerial = GetFetchXmlForActiveInventSerials();


                    //retrieve result for invent serial 
                    var resultforInventSerial = Service.RetrieveAll(orgService, new FetchExpression(fetchXmlforInventSerial));

                    Console.WriteLine("Total " + resultforInventSerial.Count + " Invent Serial Records were fetched from CRM");

                    // Iterate through each entity in the result and update 'a2c_name'
                    foreach (var entity in resultforInventSerial)
                    {
                        // Copy 'a2c_acinvent_inventserialid' value to 'a2c_name'
                        var recordId = entity.GetAttributeValue<Guid>("a2c_acinventserialid");
                        var inventSerialId = entity.GetAttributeValue<String>("a2c_acinvent_inventserialid");
                        entity["a2c_name"] = inventSerialId.ToString(); // Assuming 'a2c_name' is a string field

                        // Update the entity
                        orgService.Update(entity);
                        count++;
                        Console.WriteLine(count + " : " + inventSerialId + " udpated into autonumber for record with id " + recordId);
                    }




                    // Retrieve the existing records
                    var fetchExpression = new FetchExpression(fetchXmlforAccountID);
                    //var resultforAccountID = Service.RetrieveAll(orgService, new FetchExpression(fetchXmlforAccountID));
                    var resultforKYC = Service.RetrieveAll(orgService, new FetchExpression(fetchXmlforKYCRequired));

                    int counter = 1;

                    // Update each record with the auto number
                    //foreach (var record in resultforAccountID)
                    //{
                    //    // Generate the auto number value (e.g., ACC-0000001)
                    //    string autoNumber = $"ACC-{counter.ToString("D7")}";

                    //    // Update the record with the auto number
                    //    var account = new Entity("account")
                    //    {
                    //        Id = record.Id,
                    //    };
                    //    account["a2c_crmaccountid"] = autoNumber;

                    //    orgService.Update(account);

                    //    counter++;
                    //}

                    foreach (var record in resultforKYC)
                    {
                        // Update the record with the auto number
                        var account = new Entity("account")
                        {
                            Id = record.Id,
                        };
                        account["a2c_kycrequired"] = true;
                        OptionSetValue newValue = new OptionSetValue(602780000);
                        account["a2c_watchlisttype"] = newValue;

                        orgService.Update(account);

                        counter++;
                    }


                    // Update each record with the auto number
                    //foreach (var record in resultforAccountID)
                    //{
                    //    // Generate the auto number value (e.g., ACC-0000001)
                    //    string autoNumber = $"ACC-{counter.ToString("D7")}";

                    //    // Update the record with the auto number
                    //    var account = new Entity("account")
                    //    {
                    //        Id = record.Id,
                    //    };
                    //    account["a2c_crmaccountid"] = autoNumber;

                    //    orgService.Update(account);

                    //    counter++;
                    //}



                    var result = Service.RetrieveAll(orgService, new FetchExpression(fetchXml));
                    count = result.Count();
                    Console.WriteLine("Total " + count + " Active End User accounts were fetched from CRM");
                    List<Account> accounts = new List<Account>();

                    foreach (var record in result)
                    {
                        Account account = new Account();
                        if (record.Attributes.Contains("name"))
                        {
                            account.Name = record.Attributes["name"].ToString();
                        }
                        if (record.Attributes.Contains("accountid"))
                        {
                            account.AccountId = record.Attributes["accountid"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_accountsalesstage"))
                        {
                            account.A2C_AccountSalesStage = record.FormattedValues["a2c_accountsalesstage"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_accounttype"))
                        {
                            account.A2C_AccountType = record.FormattedValues["a2c_accounttype"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_zoominfoaccountid"))
                        {
                            account.A2C_ZoomInfoAccountId = record.Attributes["a2c_zoominfoaccountid"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_employeeband"))
                        {
                            account.A2C_EmployeeBand = record.FormattedValues["a2c_employeeband"].ToString();
                        }
                        if (record.Attributes.Contains("mkt_accountdomain"))
                        {
                            account.A2C_Domain = record.Attributes["mkt_accountdomain"].ToString();
                        }
                        if (record.Attributes.Contains("ownerid"))
                        {
                            EntityReference refOwnerName = (EntityReference)record.Attributes["ownerid"];
                            string ownerName = refOwnerName.Name;
                            Guid ownerId = refOwnerName.Id;
                            account.Owner = ownerName;
                            account.OwnerId = ownerId;
                        }
                        if (record.Attributes.Contains("new_country") && record.FormattedValues["new_country"] != null)
                        {
                            account.Country = record.FormattedValues["new_country"].ToString();
                        }
                        if (record.Attributes.Contains("modifiedon"))
                        {
                            account.Lastmodified = record.Attributes["modifiedon"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_accountpriority"))
                        {
                            account.AccountPriority = record.FormattedValues["a2c_accountpriority"].ToString();
                        }
                        if (record.Attributes.Contains("a2c_accountpotential"))
                        {
                            account.AccountPotential = record.FormattedValues["a2c_accountpotential"].ToString();
                        }
                        if (record.Attributes.Contains("primarycontactid"))
                        {
                            EntityReference refContactName = (EntityReference)record.Attributes["primarycontactid"];
                            string primaryContactName = refContactName.Name;
                            account.PrimaryContact = primaryContactName;
                            Guid contactId = refContactName.Id;
                            account.PrimaryContactId = contactId;
                        }
                        if (record.Attributes.Contains("a2c_schedulefollowup"))
                        {
                            account.ScheduledFollowup = record.Attributes["a2c_schedulefollowup"].ToString();
                        }

                        accounts.Add(account);
                    }

                    List<Account> duplicateAccounts = (from aacc in accounts
                                                       join ad in (from a in accounts
                                                                   group a by a.Name into g
                                                                   where g.Count() > 1
                                                                   select new { Name = g.Key, Count = g.Count() }) on aacc.Name equals ad.Name
                                                       where aacc.A2C_AccountType == "End User"
                                                       orderby aacc.Name, aacc.AccountId
                                                       select aacc).ToList();
                    count = duplicateAccounts.Count();
                    Console.WriteLine("Total " + count + " Duplicates were found");

                    var accountGroups = duplicateAccounts.GroupBy(a => a.Name);

                    List<Account> groupsWithParentChild = new List<Account>();
                    List<Account> groupsWithoutParentChild = new List<Account>();
                    List<Account> deactivatedAccounts = new List<Account>();

                    foreach (var group in accountGroups)
                    {
                        bool allNotStartedorNull = group.All(a => a.A2C_AccountSalesStage == "01. Not started" || string.IsNullOrEmpty(a.A2C_AccountSalesStage));

                        if (allNotStartedorNull)
                        {
                            var parentAccount = group.FirstOrDefault(a => !string.IsNullOrEmpty(a.A2C_ZoomInfoAccountId))
                                ?? group.FirstOrDefault(a => !string.IsNullOrEmpty(a.A2C_EmployeeBand))
                                ?? group.FirstOrDefault(a => !string.IsNullOrEmpty(a.A2C_Domain))
                                ?? group.FirstOrDefault();

                            parentAccount.ParentChild = "Parent";

                            foreach (var account in group.Where(a => a != parentAccount))
                            {
                                account.ParentChild = "Child";
                            }
                        }
                        else
                        {
                            var countries = group.Select(a => a.Country).Where(c => !string.IsNullOrEmpty(c)).Distinct().ToList();
                            if (countries.Count > 1) // Check if there are different countries within the group
                            {
                                // Leave the entire group untagged
                            }
                            else
                            {
                                var highestPriorityStage = group.Max(a => GetAccountSalesStagePriority(a.A2C_AccountSalesStage));
                                var accountsWithHighestPriority = group.Where(a => GetAccountSalesStagePriority(a.A2C_AccountSalesStage) == highestPriorityStage).ToList();

                                if (accountsWithHighestPriority.Count == 1)
                                {
                                    var parentAccount = accountsWithHighestPriority.First();
                                    parentAccount.ParentChild = "Parent";

                                    foreach (var account in group.Where(a => a != parentAccount))
                                    {
                                        account.ParentChild = "Child";
                                    }
                                }
                                // If multiple accounts have the highest priority stage, do nothing
                            }
                        }

                        if (group.Any(a => a.ParentChild == "Parent"))
                        {
                            groupsWithParentChild.AddRange(group);
                        }
                        else
                        {
                            groupsWithoutParentChild.AddRange(group);
                        }
                    }

                    var taggedGroups = groupsWithParentChild.GroupBy(a => a.Name);
                    foreach (var group in taggedGroups)
                    {
                        var childAccounts = group.Where(a => a.ParentChild != "Parent").ToList();
                        var parAccount = group.FirstOrDefault(a => a.ParentChild == "Parent");
                        var childAccountsWithPrimaryContact = childAccounts.Where(a => !string.IsNullOrEmpty(a.PrimaryContact)).ToList();
                        var childAccountsWithOtherOwner = childAccounts.Where(a => a.Owner != "# CRM Admin").ToList();

                        // Check if the parent account has a primary contact
                        if (parAccount != null && string.IsNullOrEmpty(parAccount.PrimaryContact) && childAccountsWithPrimaryContact.Count > 0)
                        {
                            var primaryContact = childAccountsWithPrimaryContact.First();
                            parAccount.PrimaryContact = primaryContact.PrimaryContact;

                            // Update the primary contact of the parent account in CRM
                            Entity accountToUpdate = new Entity("account");
                            accountToUpdate.Id = new Guid(parAccount.AccountId);
                            accountToUpdate["primarycontactid"] = new EntityReference("contact", primaryContact.PrimaryContactId);

                            orgService.Update(accountToUpdate);
                        }

                        if (parAccount != null && parAccount.Owner == "# CRM Admin" && childAccountsWithOtherOwner.Count > 0)
                        {
                            var childAccountWithOtherOwner = childAccountsWithOtherOwner.First();
                            parAccount.Owner = childAccountWithOtherOwner.Owner;

                            // Update the owner of the parent account in CRM
                            Entity accountToUpdate = new Entity("account");
                            accountToUpdate.Id = new Guid(parAccount.AccountId);
                            if (childAccountWithOtherOwner.OwnerId != new Guid("24603699-8040-eb11-bf68-000d3a874495"))
                            {
                                accountToUpdate["ownerid"] = new EntityReference("systemuser", childAccountWithOtherOwner.OwnerId);
                            }
                            else
                            {
                                accountToUpdate["ownerid"] = new EntityReference("systemuser", new Guid("75514c28-a38d-ed11-81ad-0022481b5481"));
                            }


                            orgService.Update(accountToUpdate);
                        }

                        foreach (var childAccount in childAccounts)
                        {
                            // Retrieve contacts associated with the child account
                            QueryExpression contactQuery = new QueryExpression("contact");
                            contactQuery.Criteria.AddCondition("parentcustomerid", ConditionOperator.Equal, childAccount.AccountId);
                            EntityCollection contacts = orgService.RetrieveMultiple(contactQuery);

                            // Attach contacts to the parent account
                            foreach (var contact in contacts.Entities)
                            {
                                EntityReference contactRef = contact.ToEntityReference();
                                var parentAccount = group.FirstOrDefault(a => a.ParentChild == "Parent");
                                if (parentAccount != null)
                                {
                                    orgService.Associate("account", Guid.Parse(parentAccount.AccountId), new Relationship("contact_customer_accounts"), new EntityReferenceCollection() { contactRef });
                                }
                            }

                            // Retrieve opportunities associated with the child account
                            QueryExpression opportunityQuery = new QueryExpression("opportunity");
                            opportunityQuery.Criteria.AddCondition("parentaccountid", ConditionOperator.Equal, childAccount.AccountId);
                            EntityCollection opportunities = orgService.RetrieveMultiple(opportunityQuery);

                            // Attach opportunities to the parent account
                            foreach (var opportunity in opportunities.Entities)
                            {
                                EntityReference opportunityRef = opportunity.ToEntityReference();
                                var parentAccount = group.FirstOrDefault(a => a.ParentChild == "Parent");
                                if (parentAccount != null)
                                {
                                    orgService.Associate("account", Guid.Parse(parentAccount.AccountId), new Relationship("opportunity_customer_accounts"), new EntityReferenceCollection() { opportunityRef });
                                }
                            }

                            // Retrieve leads associated with the child account
                            QueryExpression leadQuery = new QueryExpression("lead");
                            leadQuery.Criteria.AddCondition("parentaccountid", ConditionOperator.Equal, childAccount.AccountId);
                            EntityCollection leads = orgService.RetrieveMultiple(leadQuery);

                            // Attach leads to the parent account
                            foreach (var lead in leads.Entities)
                            {
                                EntityReference leadRef = lead.ToEntityReference();
                                var parentAccount = group.FirstOrDefault(a => a.ParentChild == "Parent");
                                if (parentAccount != null)
                                {
                                    orgService.Associate("account", Guid.Parse(parentAccount.AccountId), new Relationship("lead_parent_account"), new EntityReferenceCollection() { leadRef });
                                }
                            }

                            // Deactivate child accounts
                            orgService.Execute(new SetStateRequest
                            {
                                EntityMoniker = new EntityReference("account", Guid.Parse(childAccount.AccountId)),
                                State = new OptionSetValue(1), // 1 = Inactive
                                Status = new OptionSetValue(2) // 2 = Inactive
                            });
                            deactivatedAccounts.Add(childAccount);
                        }
                    }

                    Console.WriteLine("Microsoft Dynamics CRM version {0}.", ((RetrieveVersionResponse)orgService.Execute(new RetrieveVersionRequest())).Version);
                }
            }
            catch (FaultException<OrganizationServiceFault> osFaultException)
            {
                Console.WriteLine("Fault Exception caught");
                Console.WriteLine(osFaultException.Detail.Message);
            }
            catch (Exception e)
            {
                Console.WriteLine("Uncaught Exception");
                Console.WriteLine(e);
            }
        }

        static int GetAccountSalesStagePriority(string stage)
        {
            switch (stage)
            {
                case "02. Contact started":
                    return 3;
                case "03. Contact Established":
                    return 4;
                case "04. Nurture > Engaging & Profiling":
                    return 5;
                case "05. Nurture > Reviewing Sample":
                    return 6;
                case "06. Nurture > On Hold":
                    return 7;
                case "07. Qualified Sales Lead":
                    return 8;
                case "08. Opportunity":
                    return 9;
                case "09. Customer":
                    return 10;
                case "10. Parked Permanently":
                    return 2;
                default:
                    return 0;
            }
        }

        private static string GetConnectionString()
        {
            return ConfigurationManager.ConnectionStrings["CRMConnectionString"].ConnectionString;
        }

        static string GetFetchXml()
        {
            return @"
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                    <entity name='account'>
                        <attribute name='accountid' />
                        <attribute name='name' />
                        <attribute name='a2c_accounttype' />
                        <attribute name='a2c_accountsalesstage' />
                        <attribute name='ownerid' />
                        <attribute name='new_country' />
                        <attribute name='modifiedon' />
                        <attribute name='a2c_accountpriority' />
                        <attribute name='a2c_accountpotential' />
                        <attribute name='primarycontactid' />
                        <attribute name='a2c_schedulefollowup' />
                        <!-- Add other attributes here -->
                        <filter>
                            <condition attribute='a2c_accounttype' operator='eq' value='602780002' />
                            <condition attribute='statecode' operator='eq' value='0' />
                        </filter>
                    </entity>
                </fetch>";
        }

        static string GetFetchXmlForAccountID()
        {
            return @"
        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='account'>
                <attribute name='accountid' />
                <attribute name='createdon' />
                <order attribute='createdon' descending='false' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                </filter>
            </entity>
        </fetch>";
        }

        static string GetFetchXmlForKYCRequired()
        {
            return @"
        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='account'>
                <attribute name='name' />
                <attribute name='a2c_kycrequired' />
                <filter type='and'>
                    <condition attribute='statecode' operator='eq' value='0' />
                    <condition attribute='a2c_kycrequired' operator='ne' value='1' />
                </filter>
            </entity>
        </fetch>";
        }


        static string GetFetchXmlForActiveInventSerials()
        {
            return @"
    <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
        <entity name='a2c_acinventserial'>
            <attribute name='a2c_name' />
            <attribute name='a2c_acinvent_inventserialid' />
            <filter type='and'>
                <condition attribute='statecode' operator='eq' value='0' />
                <condition attribute='a2c_acinvent_inventserialid' operator='not-null' />
                <condition attribute='a2c_name' operator='null' /> 
                <condition attribute='createdon' operator='on-or-after' value='2021-01-01' /> 
            </filter>
        </entity>
    </fetch>";
        }

        public static string GetFetchXmlForDuplicateInventSerial()
        {
            string fetchXml = @"
        <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
            <entity name='a2c_acinventserial'>
                <attribute name='a2c_acinventserialid' />
                <attribute name='a2c_acinvent_inventserialid' />
                <attribute name='a2c_acinvent_areaid' />
                <attribute name='a2c_acinvent_itemid_txt' />
                <attribute name='a2c_acinvent_receiveddate' />
                <attribute name='a2c_acinvent_ponumber' />
                <attribute name='a2c_salesordernumbername' />
                <attribute name='a2c_acinvent_status_txt' />
                <attribute name='statuscodename' />
                <link-entity name='a2c_acinventserial' from='a2c_acinvent_inventserialid' to='a2c_acinvent_inventserialid' alias='b' link-type='inner'>
                    <attribute name='a2c_acinvent_inventserialid' alias='b_a2c_acinvent_inventserialid' />
                    <filter type='group' >
                        <condition attribute='a2c_acinvent_inventserialid' operator='not-null' />
                        <condition attribute='a2c_acinvent_inventserialid' operator='not-like' value='' />
                    </filter>
                    <attribute name='a2c_acinvent_inventserialid' />
                    <filter type='and'>
                        <condition attribute='statecode' operator='eq' value='0' />
                    </filter>
                    <link-entity name='a2c_acinventserial' from='a2c_acinvent_inventserialid' to='a2c_acinvent_inventserialid' alias='a' link-type='inner'>
                        <attribute name='a2c_acinvent_inventserialid' alias='a_a2c_acinvent_inventserialid' />
                        <filter type='and'>
                            <condition attribute='statecode' operator='eq' value='0' />
                            <condition attribute='statuscodename' operator='eq' value='Active' />
                        </filter>
                    </link-entity>
                </link-entity>
                <order attribute='a2c_acinvent_inventserialid' />
                <order attribute='a2c_acinvent_areaid' />
                <order attribute='a2c_acinvent_itemid_txt' />
            </entity>
        </fetch>";

            return fetchXml;
        }


        static void DuplicateAcInvent()
        {
            Console.WriteLine(" ========== De-Duplication Utility ============ ");
            Console.WriteLine("\n");
            Console.WriteLine(" ---------------------------------------------- ");
            Console.WriteLine("\n");

            Console.WriteLine("Connecting with A2C Production CRM");
            Console.WriteLine("\n");
            string connectionString = GetConnectionString();
            List<Entity> recordsToDeactivate = new List<Entity>();
            List<Entity> allRecords = new List<Entity>();
            using (CrmServiceClient conn = new CrmServiceClient(connectionString))
            {
                Console.WriteLine("Connection successful");
                Console.WriteLine("\n");
                Console.WriteLine("Fetching Results");
                Console.WriteLine("\n");
                int count = 0;



                // Cast the proxy client to the IOrganizationService interface.
                IOrganizationService orgService = (IOrganizationService)conn.OrganizationWebProxyClient ?? conn.OrganizationServiceProxy;
                // Define the query expression
                QueryExpression query = new QueryExpression("a2c_acinventserial");
                query.ColumnSet = new ColumnSet("a2c_acinventserialid", "a2c_acinvent_inventserialid", "a2c_acinvent_receiveddate", "a2c_acinvent_areaid", "a2c_acinvent_itemid_txt", "a2c_acinvent_status_txt");
                query.Criteria.AddCondition("statuscodename", ConditionOperator.Equal, "Active");
                query.Criteria.AddCondition(new ConditionExpression("a2c_acinvent_areaid", ConditionOperator.NotNull));
                query.Criteria.AddCondition(new ConditionExpression("a2c_acinvent_areaid", ConditionOperator.NotNull));
                query.PageInfo = new PagingInfo();
                query.PageInfo.PageNumber = 1;
                query.PageInfo.Count = 5000;

                while (true)
                {
                    // Retrieve records
                    EntityCollection results = orgService.RetrieveMultiple(query);

                    // Add retrieved records to the list
                    foreach (var record in results.Entities)
                    {
                        allRecords.Add(record);
                    }

                    // Check if there are more records
                    if (!results.MoreRecords)
                    {
                        break;
                    }

                    // Increment the page number
                    query.PageInfo.PageNumber++;
                    // Set the paging cookie to get the next set of records
                    query.PageInfo.PagingCookie = results.PagingCookie;
                }

                Console.WriteLine($"Total records fetched: {allRecords.Count}");

                // Dictionary to hold groups of records based on invent serial id
                Dictionary<string, List<Entity>> groups = new Dictionary<string, List<Entity>>();


                // Group records using LINQ
                var groupedRecords = allRecords
                    .GroupBy(record => new
                    {
                        InventSerialId = record.GetAttributeValue<string>("a2c_acinvent_inventserialid"),
                        AreaId = record.GetAttributeValue<string>("a2c_acinvent_areaid").ToUpper(),
                        ItemId = record.GetAttributeValue<string>("a2c_acinvent_itemid_txt")
                    })
                     .Where(group => group.Count() > 1)
                    .OrderBy(group => group.Key.InventSerialId)
                    .Select(group => group.OrderByDescending(record => record.GetAttributeValue<DateTime>("a2c_acinvent_receiveddate")).ToList())
                    .ToList();

                ExportToExcel(groupedRecords);

                int totalUpdatedRecords = 0;

                using (StreamWriter writer = new StreamWriter("DeactivationLog1.txt"))
                {
                    foreach (var group in groupedRecords)
                    {
                        // Keep the first record as it's the latest one
                        var latestRecord = group.First();

                        // Deactivate all other records in the group
                        foreach (var record in group.Skip(1))
                        {
                            record["statecode"] = new OptionSetValue(1); // Assuming 'Inactive' status is mapped to value 1
                            record["statuscode"] = new OptionSetValue(2);
                            orgService.Update(record);
                            totalUpdatedRecords++;

                            // Write information about the deactivated record to the text file
                            writer.WriteLine($"Record deactivated - ID: {record.GetAttributeValue<Guid>("a2c_acinventserialid")}, Invent Serial ID: {record.GetAttributeValue<string>("a2c_acinvent_inventserialid")}, Area ID: {record.GetAttributeValue<string>("a2c_acinvent_areaid")}, Item ID: {record.GetAttributeValue<string>("a2c_acinvent_itemid_txt")}");
                        }
                        //if (group.Count > 1)
                        //{
                        //    // Mark non-latest records for deactivation
                        //    for (int i = 1; i < group.Count; i++)
                        //    {
                        //        group[i]["statecode"] = new OptionSetValue(1);
                        //        group[i]["statuscode"] = new OptionSetValue(2);
                        //        orgService.Update(group[i]);
                        //    }
                        //}
                    }
                }
                


                Console.WriteLine("De-duplication and deactivation completed.");





            }

        }


        // Method to export records to Excel
        static void ExportToExcel(List<List<Entity>> groupedRecords)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Records");

                // Add headers
                worksheet.Cells[1, 1].Value = "Invent Serial ID";
                worksheet.Cells[1, 2].Value = "Area ID";
                worksheet.Cells[1, 3].Value = "Item ID";
                worksheet.Cells[1, 4].Value = "Received Date";
                worksheet.Cells[1, 5].Value = "Status";

                // Populate data
                int row = 2;
                foreach (var group in groupedRecords)
                {
                    foreach (var record in group)
                    {
                        worksheet.Cells[row, 1].Value = record.GetAttributeValue<string>("a2c_acinvent_inventserialid");
                        worksheet.Cells[row, 2].Value = record.GetAttributeValue<string>("a2c_acinvent_areaid");
                        worksheet.Cells[row, 3].Value = record.GetAttributeValue<string>("a2c_acinvent_itemid_txt");
                        worksheet.Cells[row, 4].Value = record.GetAttributeValue<DateTime>("a2c_acinvent_receiveddate");
                        worksheet.Cells[row, 5].Value = record.GetAttributeValue<string>("a2c_acinvent_status_txt");
                        row++;
                    }
                }

                // Save Excel package
                FileInfo excelFile = new FileInfo("Records1.xlsx");
                excelPackage.SaveAs(excelFile);
            }

            Console.WriteLine("Records exported to Excel.");
        }






    }
}

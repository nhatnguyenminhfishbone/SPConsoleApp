// See https://aka.ms/new-console-template for more information
//using Microsoft.SharePoint.Client;
//using System.Collections.Generic;

//using (ClientContext clientContext = new ClientContext("https://omniadevcloud.sharepoint.com/sites/nhatnm"))
//{
//     clientcontext.Web.Lists.GetById - This option also can be used to get the list using List GUID
//     This value is NOT List internal name
//    List targetList = clientContext.Web.Lists.GetByTitle("List create from code");

//    ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
//    ListItem oItem = targetList.AddItem(oListItemCreationInformation);
//    oItem["Title"] = "New List Item";

//    oItem.Update();
//    clientContext.ExecuteQuery();
//}

using Microsoft.SharePoint.Client;
using System.Net;
using System.Security;

//public ClientContext GetContext(Uri web, string userPrincipalName, SecureString userPassword)
//{
//    context.ExecutingWebRequest += (sender, e) =>
//    {
//        // Get an access token using your preferred approach
//        string accessToken = MyCodeToGetAnAccessToken(new Uri($"{web.Scheme}://{web.DnsSafeHost}"), userPrincipalName, new System.Net.NetworkCredential(string.Empty, userPassword).Password);
//        // Insert the access token in the request
//        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + accessToken;
//    };
//}


using CSOMDemo;
using System.Security;
using Microsoft.AspNetCore.Http.Authentication;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.VisualBasic.FileIO;
using System.Net.Mime;
using ContentType = Microsoft.SharePoint.Client.ContentType;
using System.Collections.Generic;
using Microsoft.SharePoint.News.DataModel;
using OfficeDevPnP.Core.Entities;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Microsoft.Graph;


//public static async Task Main(string[] args)
//{
Uri site = new Uri("https://02s27.sharepoint.com/CSOM%20Exercise-Nhat");
    string user = "namphuongdinh@02s27.onmicrosoft.com";
    SecureString password = GetSecureString("Nam040302@");

    // Note: The PnP Sites Core AuthenticationManager class also supports this
    using (var authenticationManager = new CSOMDemo.AuthenticationManager())
    using (var context = authenticationManager.GetContext(site, user, password))
    {
        context.Load(context.Web, p => p.Title);
        await context.ExecuteQueryAsync();
        Console.WriteLine($"Title: {context.Web.Title}");
    }
//}

static SecureString GetSecureString(string password)
{
    SecureString securePassword = new SecureString();

    foreach (char c in password)
    {
        securePassword.AppendChar(c);
    }
    return securePassword;
}

// Using CSOM create a List name “CSOM Test”
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{

//    ListCreationInformation creationInfo = new ListCreationInformation();
//    creationInfo.Title = "CSOM Test Nhat";
//    creationInfo.Description = "CSOM Test Nhat";
//    creationInfo.TemplateType = (int)ListTemplateType.GenericList;
//    // Create a new custom list    
//    List newList = context.Web.Lists.Add(creationInfo);
//    // Retrieve the custom list properties    
//    context.Load(newList);
//    // Execute the query to the server.    
//    context.ExecuteQuery();
//    // Display the custom list Title property    
//    Console.WriteLine(newList.Title);
//    Console.ReadLine();
//}

// Create term set “city-{yourname}” in dev tenant
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

//    // Get the term store by name
//    TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
//    TermGroup termGroup = termStore.Groups.GetByName("city-Nhat");
//    // Int variable - new term lcid
//    int lcid = 1033;

//    // Guid - new term guid
//    Guid newTermId = Guid.NewGuid();
//    TermSet termSet = termGroup.CreateTermSet("city-Nhatt", newTermId, lcid);
//    context.ExecuteQuery();
//}

// Create 2 terms “Ho Chi Minh” and “Stockholm” in termset “city-{yourname}”
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(context);

//    // Get the term store by name
//    TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
//    TermGroup termGroup = termStore.Groups.GetByName("city-Nhat");
//    // Int variable - new term lcid
//    int lcid = 1033;

//    // Guid - new term guid
//    Guid newTermId1 = Guid.NewGuid();
//    Guid newTermId2 = Guid.NewGuid();
//    TermSet termSet = termGroup.TermSets.GetByName("city-Nhatt");
//    string termName1 = "Ho Chi Minh";
//    string termName2 = "Stockholm";

//    Term newTerm1 = termSet.CreateTerm(termName1, lcid, newTermId1);
//    Term newTerm2 = termSet.CreateTerm(termName2, lcid, newTermId2);
//    context.ExecuteQuery();
//}

// Create site fields “about” type text and field “city” type taxonomy
// Create site content type “CSOM Test content type” then add this content type to list “CSOM test”, add fields “about” and “city” to this content type.
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    Web web = context.Web;
//    ContentType contentTypes = web.ContentTypes.GetById("0x01");
//    ContentTypeCreationInformation newContentType = new ContentTypeCreationInformation
//    {
//        Name = "CSOM Test content type by Nhat test",
//        Group = "List Content Types",
//        Description = "Custom content type based on the built-in Item content type.",
//        ParentContentType = contentTypes
//    };                 // Add the new content type to the site collection
//    ContentType contentType = web.ContentTypes.Add(newContentType);
//    context.Load(contentType);
//    context.ExecuteQuery();
//    // Add new field
//    FieldCollection fields = web.Fields;
//    context.Load(fields);
//    context.ExecuteQuery(); string fieldXml = "<Field DisplayName='AboutNhatt' Name='AboutNhatt' Type='Text' />"; Field field = fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
//    context.Load(field);
//    context.ExecuteQuery();
//    string cityFieldXml = "<Field DisplayName='CityNhatt' Name='CityNhatt' Type='TaxonomyFieldTypeMulti' />";
//    Field cityField = fields.AddFieldAsXml(cityFieldXml, true, AddFieldOptions.DefaultValue);
//    context.Load(cityField);
//    context.ExecuteQuery(); // Associate the "City" field with a term set
//    Guid termStoreId = Guid.Empty;
//    Guid termSetId = Guid.Empty;
//    GetTaxonomyFieldInfo(context, out termStoreId, out termSetId);  // Retrieve the "City" field as a TaxonomyField
//    TaxonomyField taxonomyField = context.CastTo<TaxonomyField>(cityField);
//    taxonomyField.SspId = termStoreId;
//    taxonomyField.TermSetId = termSetId;
//    taxonomyField.TargetTemplate = String.Empty;
//    taxonomyField.AnchorId = Guid.Empty;
//    taxonomyField.Update();
//    context.ExecuteQuery();
//    // Add the new field to the content type
//    FieldLinkCreationInformation fieldLink = new FieldLinkCreationInformation
//    {
//        Field = field
//    };
//    contentType.FieldLinks.Add(fieldLink);
//    FieldLinkCreationInformation cityFieldLink = new FieldLinkCreationInformation
//    {
//        Field = cityField
//    };
//    contentType.FieldLinks.Add(cityFieldLink);
//    contentType.Update(true);
//    context.ExecuteQuery();
//    // Add the new content type to the "Custom List" list
//    /* List customList = web.Lists.GetByTitle("CSOM Exercise");*/
//    List list = context.Web.Lists.GetByTitle("CSOM Test Nhat");
//    list.ContentTypesEnabled = true;
//    list.ContentTypes.AddExistingContentType(contentType);
//    list.Update();
//    context.ExecuteQuery();
//}

//void GetTaxonomyFieldInfo(ClientContext clientContext, out Guid termStoreId, out Guid termSetId)
//{
//    termStoreId = Guid.Empty;
//    termSetId = Guid.Empty; TaxonomySession session = TaxonomySession.GetTaxonomySession(clientContext);
//    TermStore termStore = session.GetDefaultSiteCollectionTermStore();
//    TermSetCollection termSets = termStore.GetTermSetsByName("City-Nhatt", 1033); clientContext.Load(termSets, tsc => tsc.Include(ts => ts.Id));
//    clientContext.Load(termStore, ts => ts.Id);
//    clientContext.ExecuteQuery(); termStoreId = termStore.Id;
//    termSetId = termSets.FirstOrDefault().Id;
//}

// In list “CSOM test” set “CSOM Test content type Nhat” as default content type
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    List list = context.Web.Lists.GetByTitle("CSOM test Nhat");
//    ContentTypeCollection currentCtOrder = list.ContentTypes;
//    context.Load(currentCtOrder);
//    context.ExecuteQuery();

//    IList<ContentTypeId> reverceOrder = new List<ContentTypeId>();
//    foreach (ContentType ct in currentCtOrder)
//    {
//        if (ct.Name.Equals("CSOM Test content type by Nhat test"))
//        {
//            reverceOrder.Add(ct.Id);
//        }
//    }
//    list.RootFolder.UniqueContentTypeOrder = reverceOrder;
//    list.RootFolder.Update();
//    list.Update();
//    context.ExecuteQuery();
//    Console.WriteLine("Set default content type successfully");
//}

// Create 5 list items to list with some value in field “about” and “city”
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    List targetList = context.Web.Lists.GetByTitle("CSOM Test Nhat");

//    Random random = new Random();
//    for (int i = 0; i < 5; i++)
//    {
//        ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
//        ListItem oItem = targetList.AddItem(oListItemCreationInformation);
//        oItem["AboutNhatt"] = "Random AboutNhat column " + i;
//        oItem.Update();
//        context.ExecuteQuery();

//        var field = context.Web.Fields.GetByTitle("CityNhatt");
//        context.Load(field);
//        await context.ExecuteQueryAsync();
//        var taxField = context.CastTo<TaxonomyField>(field);
//        int randomNumber = random.Next(1, 3);
//        if (randomNumber == 1) {
//            taxField.SetFieldValueByValue(oItem, new TaxonomyFieldValue()
//            {
//                WssId = -1, // alway let it -1
//                Label = "Stockholm",
//                TermGuid = "734d30e7-ff16-4ec1-9c91-379acd6773e0"
//            });
//        }
//        else
//        {
//            taxField.SetFieldValueByValue(oItem, new TaxonomyFieldValue()
//            {
//                WssId = -1, // alway let it -1
//                Label = "Ho Chi Minh",
//                TermGuid = "e2a40d15-a211-45c0-8f15-d7efbf9df08c"
//            });
//        }

//        oItem.Update();
//        context.ExecuteQuery();
//    }

//    Console.WriteLine("Create 5 list items successfully");
//}

// Update site field “about” set default value for it to “about default” then create 2 new list items.
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    List targetList = context.Web.Lists.GetByTitle("CSOM Test Nhat");

//    // Get field from list using internal name or display name
//    Field oField = targetList.Fields.GetByInternalNameOrTitle("AboutNhatt");

//    // Set field default value
//    oField.DefaultValue = "about default";

//    oField.Update();
//    context.ExecuteQuery();

//    // Get the list "CSOM Test Nhat"
//    context.Load(targetList);
//    await context.ExecuteQueryAsync();

//    // Create two new list items with random values
//    Random random = new Random();
//    for (int i = 0; i < 2; i++)
//    {
//        ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
//        ListItem oItem = targetList.AddItem(oListItemCreationInformation);
//        var field = context.Web.Fields.GetByTitle("CityNhatt");
//        context.Load(field);
//        await context.ExecuteQueryAsync();
//        var taxField = context.CastTo<TaxonomyField>(field);
//        int randomNumber = random.Next(1, 3);
//        if (randomNumber == 1)
//        {
//            taxField.SetFieldValueByValue(oItem, new TaxonomyFieldValue()
//            {
//                WssId = -1, // alway let it -1
//                Label = "Stockholm",
//                TermGuid = "734d30e7-ff16-4ec1-9c91-379acd6773e0"
//            });
//        }
//        else
//        {
//            taxField.SetFieldValueByValue(oItem, new TaxonomyFieldValue()
//            {
//                WssId = -1, // alway let it -1
//                Label = "Ho Chi Minh",
//                TermGuid = "e2a40d15-a211-45c0-8f15-d7efbf9df08c"
//            });
//        }
//        oItem.Update();
//        await context.ExecuteQueryAsync();
//    }

//    Console.WriteLine("Update site field and create 2 list items successfully");
//}

// Update site field “city” set default value for it to “Ho Chi Minh” then create 2 new list items.
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    List targetList = context.Web.Lists.GetByTitle("CSOM Test Nhat");
//    Microsoft.SharePoint.Client.Field oField = targetList.Fields.GetByInternalNameOrTitle("CityNhatt");
//    ListItemCreationInformation oListItemCreationInformation = new ListItemCreationInformation();
//    TaxonomyField taxColumn = context.CastTo<TaxonomyField>(context.Web.Fields.GetByTitle("CityNhatt"));

//    context.Load(taxColumn);
//    context.ExecuteQuery();
//     initialize taxonomy field value
//    var defaultValue = new TaxonomyFieldValue()
//    {
//        WssId = -1, // alway let it -1
//        Label = "Ho Chi Minh",
//        TermGuid = "e2a40d15-a211-45c0-8f15-d7efbf9df08c"
//    };
//    retrieve validated taxonomy field value
//    var validatedValue = taxColumn.GetValidatedString(defaultValue);
//    context.ExecuteQuery();
//    set default value for a taxonomy field
//    taxColumn.DefaultValue = validatedValue.Value;

//    taxColumn.UpdateAndPushChanges(true);
//    context.ExecuteQuery();

//     Create two new list items with random values
//    for (int i = 0; i < 2; i++)
//    {
//        ListItem oItem = targetList.AddItem(oListItemCreationInformation);
//        oItem["Title"] = "Random test default value CityNhat" + i;
//        oItem.Update(); 
//        context.ExecuteQuery();
//    }

//    Console.WriteLine("Update site field 'city' and create 2 new list items successfully");
//}


// Write CAML query to get list items where field “about” is not “about default”
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    List targetList = context.Web.Lists.GetByTitle("CSOM Test Nhat");
//    string camlQuery = @"<View>
//                        <Query>
//                            <Where>
//                                <Neq>
//                                    <FieldRef Name='AboutNhatt' />
//                                    <Value Type='Text'>about default</Value>
//                                </Neq>
//                            </Where>
//                        </Query>
//                    </View>";

//    ListItemCollection items = targetList.GetItems(new CamlQuery { ViewXml = camlQuery });
//    context.Load(items);
//    context.ExecuteQuery();

//    foreach (var item in items)
//    {
//        Console.WriteLine(item["AboutNhatt"].ToString());
//    }
//}


// Create List View by CSOM order item newest in top and only show list item where “city” field has value “Ho Chi Minh”, View Fields: Id, Name, City, About
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    var list = context.Web.Lists.GetByTitle("CSOM Test Nhat");
//    context.Load(list);
//    context.ExecuteQuery();

//    // Create a new view
//    var viewCreationInformation = new ViewCreationInformation()
//    {
//        Title = "New view",
//        RowLimit = 10,
//        ViewFields = new string[] { "ID", "Title", "CityNhatt", "AboutNhatt" },
//        Query = @"<OrderBy><FieldRef Name='Modified' Ascending='FALSE' /></OrderBy>
//                                              <Where>
//                                                 <Eq>
//                                                    <FieldRef Name='CityNhatt' />
//                                                    <Value Type='TaxonomyFieldType'>Ho Chi Minh</Value>
//                                                 </Eq>
//                                              </Where>"
//    };

//    var view = list.Views.Add(viewCreationInformation);
//    view.DefaultView = false;
//    view.Update();
//    context.ExecuteQuery();
//}

// Write function update list items in batch, try to update 2 items every time and update field “about” which have value “about default” to “Update script”. (CAML)
//using (var authenticationManager = new CSOMDemo.AuthenticationManager())
//using (var context = authenticationManager.GetContext(site, user, password))
//{
//    var list = context.Web.Lists.GetByTitle("CSOM Test Nhat");
//    var query = new CamlQuery
//    {
//        ViewXml = "<View><Query><Where><Eq><FieldRef Name='AboutNhatt'/><Value Type='Text'>about default</Value></Eq></Where></Query><RowLimit>2</RowLimit></View>"
//    };
//    var items = list.GetItems(query);
//    context.Load(items);
//    await context.ExecuteQueryAsync();
//    for (int i = 0; i < items.Count; i++)
//    {
//        items[i]["AboutNhatt"] = "Update script";
//        items[i].Update();
//    }
//    items = list.GetItems(query);
//    context.Load(items);
//    await context.ExecuteQueryAsync();

//}

// Create new field “author” type people in list “CSOM Test” then migrate all list items to set user admin to field “CSOM Test Author”
using (var authenticationManager = new CSOMDemo.AuthenticationManager())
using (var context = authenticationManager.GetContext(site, user, password))
{


    var targetList = context.Web.Lists.GetByTitle("CSOM Test Nhat");

    // Create Author field
    var authorFieldXml = "<Field DisplayName='Author' Type='User' Required='FALSE' EnforceUniqueValues='FALSE' UserSelectionMode='PeopleOnly' UserSelectionScope='0' />";
    var authorField = targetList.Fields.AddFieldAsXml(authorFieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
    authorField.ReadOnlyField = false;
    context.ExecuteQuery();

    // Set user "admin" as Author for all list items
    var admin = context.Web.AssociatedOwnerGroup.Users;
    var camlQuery = CamlQuery.CreateAllItemsQuery(100);
    var items = targetList.GetItems(camlQuery);
    context.Load(items);
    context.Load(authorField);
    context.ExecuteQuery();

    foreach (var item in items)
    {
        item["Author"] = admin;
        item.Update();
        context.ExecuteQuery();
    }

    Console.WriteLine("Updated Author field for all items in the list.");
}
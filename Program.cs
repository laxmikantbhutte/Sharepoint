using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Sites;
using System;
using System.Security;

namespace DemoSpo
{
    class Program
    {
        static ClientContext globalctx;
        static void Main(string[] args)
        {
            #region Variable declaration
            string tenant = "https://bhutte-admin.sharepoint.com/";
            int choice=0;
            string userName = "bhutte.laxmikant@bhutte.onmicrosoft.com";
            string passwordString = "Laxmikant@150496";
            #endregion
            TestConnect(tenant, userName, passwordString);
            while(choice!=999)
            {
            Console.WriteLine("Enter your choice");
            Console.WriteLine("\n 1.Create a Team site\n 2.Cretae a List\n 3.Add Item to a List\n 4.Add a content type to a List\n 5.Exit");
            choice = Convert.ToInt32(Console.ReadLine());
            
                switch (choice)
                {
                    case 1:
                        CreateTeamSite();
                        break;

                    case 2:
                        string siteLink;
                        Console.Write("Enter Site Link in which you want to create List - \n");
                        siteLink = Console.ReadLine();
                        Console.WriteLine("You entered", siteLink);
                        TestConnect(siteLink, userName, passwordString);
                        CreateList();
                        break;

                    case 3:
                        string addItemLink;
                        Console.Write("Enter Site Link in which you want to add item - \n");
                        addItemLink = Console.ReadLine();
                        Console.WriteLine("You entered", addItemLink);
                        TestConnect(addItemLink, userName, passwordString);
                        AddItemToList();
                        break;

                    case 4:
                        string contentTypeLink;
                        Console.Write("Enter Site Link in which you want to add content type - \n");
                        contentTypeLink = Console.ReadLine();
                        Console.WriteLine("You entered", contentTypeLink);
                        TestConnect(contentTypeLink, userName, passwordString);
                        AddContentTypeToList();
                        break;

                    case 5:
                        System.Environment.Exit(0);
                        break;

                    default:
                        Console.WriteLine("Enter Correct Choice");
                        break;

                }
            
            }

        }
        private static void TestConnect(string tenant, string userName, string passwordString)

        {
            // Get access to source site
                globalctx = new ClientContext(tenant);

                //Provide count and pwd for connecting to the source

                var passWord = new SecureString();

                foreach (char c in passwordString.ToCharArray()) passWord.AppendChar(c);

                globalctx.Credentials = new SharePointOnlineCredentials(userName, passWord);

                // Actual code for operations

                Web web = globalctx.Web;
                globalctx.Load(web);
                globalctx.ExecuteQuery();
                Console.WriteLine(string.Format("Connected to site with title of {0}", web.Title));
        }

       private static async void CreateTeamSite()
        {
            try
            {
                string siteName;
                Console.Write("Enter Site Name - \n");
                siteName = Console.ReadLine();
                Console.WriteLine("You entered", siteName);

                string description;
                Console.Write("Enter a description for site - \n");
                description = Console.ReadLine();
                Console.WriteLine("You entered ", description);

                string alis;
                Console.Write("Enter alis - \n");
                alis = Console.ReadLine();
                Console.WriteLine("You entered ", alis);

                TeamSiteCollectionCreationInformation modernteamSiteInfo = new TeamSiteCollectionCreationInformation();

                modernteamSiteInfo.DisplayName = siteName;
                modernteamSiteInfo.Description = description;
                modernteamSiteInfo.Alias = alis;
                modernteamSiteInfo.IsPublic = true;

                var createModernSite = await globalctx.CreateSiteAsync(modernteamSiteInfo);

                Console.WriteLine("site created ....");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private static void CreateList()
        {
            try
            {
                string listName;
                Console.Write("Enter List Name - \n");
                listName = Console.ReadLine();
                Console.WriteLine("You entered", listName);

                string descriptionList;
                Console.Write("Enter description of List - \n");
                descriptionList = Console.ReadLine();
                Console.WriteLine("You entered", descriptionList);

                ListCreationInformation listCreationInformation = new ListCreationInformation();
                listCreationInformation.Title = listName;
                listCreationInformation.Description = descriptionList;
                listCreationInformation.TemplateType = (int)ListTemplateType.GenericList;

                List newlist = globalctx.Web.Lists.Add(listCreationInformation);
                globalctx.Load(newlist);
                globalctx.ExecuteQuery();

                Console.WriteLine("New List Created succesfully named");
                Console.WriteLine(newlist.Title);


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        private static void AddItemToList()
        {
            try
            {
                string listNameToAddItem;
                Console.Write("Enter Name of list in which you want to add ITEM - \n");
                listNameToAddItem = Console.ReadLine();
                Console.WriteLine("You entered", listNameToAddItem);

                string itemTitle;
                Console.Write("Enter title of Item - \n");
                itemTitle = Console.ReadLine();
                Console.WriteLine("You entered", itemTitle);

                List oList = globalctx.Web.Lists.GetByTitle(listNameToAddItem);
                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(listItemCreationInformation);
                oListItem["Title"] = itemTitle;

                oListItem.Update();
                globalctx.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void AddContentTypeToList()
        {
            string contentTypename;
            Console.Write("Enter content type - \n");
            contentTypename = Console.ReadLine();
            Console.WriteLine("You entered", contentTypename);

            string contentdescription;
            Console.Write("Enter a description for content type - \n");
            contentdescription = Console.ReadLine();
            Console.WriteLine("You entered ", contentdescription);

            string groupName;
            Console.Write("Enter group name - \n");
            groupName = Console.ReadLine();
            Console.WriteLine("You entered ", groupName);

            ContentTypeCollection contentTypeColl = globalctx.Web.ContentTypes;

            ContentTypeCreationInformation contentTypeCreation = new ContentTypeCreationInformation();
            contentTypeCreation.Name = contentTypename;
            contentTypeCreation.Description = contentdescription;
            contentTypeCreation.Group = groupName;

            ContentType ct = contentTypeColl.Add(contentTypeCreation);
            globalctx.Load(ct);
            globalctx.ExecuteQuery();

            Console.WriteLine(ct.Name + " content type is created successfully");

        }

    }//cs
}//ns
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Framework.Provisioning.Connectors;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.ALM;
using System.Threading;
using System.Security;
using System.IO;
using System.Net;
using ClientSidePage = OfficeDevPnP.Core.Pages.ClientSidePage;
using OfficeDevPnP.Core.Pages;

namespace PNPProvisioningAllWebs
{
    class Program
    {
        static String O365userName, O365password;
        
        static String siteURL = "", rootsiteUrl = "", hotelTitle = "", hotelURL = "";
        static SharePointOnlineCredentials credentials;
        static String path = "";
        static String tenantUrl = "";
        static String contentypehubUrl = "/sites/contentTypeHub";


        static CookieContainer cookieContainer = new CookieContainer();
        private const String userAgentSelfContained =
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36";

        public static void Main(String[] args)
        {

            Console.WriteLine("Please provide UserName");
            O365userName = Convert.ToString(Console.ReadLine()).Trim();
            if (String.IsNullOrWhiteSpace(O365userName))
            {
                Console.WriteLine("Please provide correct UserName.");
            }
            else
            {
                //get password
                Console.WriteLine("Please provide Password");
                String pass = "";
                do
                {
                    ConsoleKeyInfo key = Console.ReadKey(true);
                    // Backspace Should Not Work
                    if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                    {
                        pass += key.KeyChar;
                        Console.Write("*");
                    }
                    else
                    {
                        if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
                        {
                            pass = pass.Substring(0, (pass.Length - 1));
                            Console.Write("\b \b");
                        }
                        else if (key.Key == ConsoleKey.Enter)
                        {
                            break;
                        }
                    }
                } while (true);
                O365password = Convert.ToString(pass).Trim();
                if (String.IsNullOrWhiteSpace(O365password))
                {
                    Console.WriteLine("Please provide correct Password.");
                }
                else
                {
                    //get app catalog url
                    Console.WriteLine();
                    Console.WriteLine("Please provide root site URL");
                    siteURL = Convert.ToString(Console.ReadLine()).Trim();
                    rootsiteUrl = siteURL;
                    if (String.IsNullOrWhiteSpace(siteURL))
                    {
                        Console.WriteLine("Please provide correct site URL.");
                    }
                    else
                    {
                        //get credential
                        credentials = GetCredentials(O365userName, O365password);

                        Console.WriteLine("Please provide hotel site Title");
                        hotelTitle = Convert.ToString(Console.ReadLine()).Trim();

                        Console.WriteLine("Please provide hotel site URL");
                        hotelURL = Convert.ToString(Console.ReadLine()).Trim();

                        createHotel();
                        //createPages();
                    }

                }

            }

            Console.WriteLine("Hit [Enter] to exit");
            Console.ReadLine();
        }

        public static void ApplyTemplateToRootSite(String WebUrl)
        {
            using (var ctx = new ClientContext(WebUrl))
            {
                ctx.Credentials = credentials;
                ctx.RequestTimeout = Timeout.Infinite;

                Web Web = ctx.Web;
                ctx.Load(Web, w => w.Title);
                ctx.ExecuteQueryRetry();

                System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(@"../Template/");

                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(rootDir.FullName + "\\pnpprovisioning", "");

                ProvisioningTemplate template = provider.GetTemplate("MyPnPProvisioningHotelBG.xml");
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();

                FileSystemConnector connector = new FileSystemConnector(rootDir.FullName + "\\pnpprovisioning", "");
                template.Connector = connector;
                try
                {
                    Console.WriteLine("Template is provisioning on " + WebUrl);
                    Web.ApplyProvisioningTemplate(template, ptai);


                }

                catch (Exception ex)
                {
                    String hotelSiteUrl = rootsiteUrl + "/" + hotelURL;
                    if (ex.Message.Contains("connection") || ex.Message.Contains("One or more"))
                    {
                        Console.WriteLine("Template is provisioning on " + WebUrl);
                        Web.ApplyProvisioningTemplate(template, ptai);

                        ApplyCommListToSite(WebUrl);

                    }
                    else
                    {
                        ApplyCommListToSite(WebUrl);
                        // Console.WriteLine("Error occured while appplying template." + ex.Message);
                    }
                }
            }
        }

        public static void ApplyTemplateToSubSite(String WebUrl, String subsitexml)
        {
            using (var ctx = new ClientContext(WebUrl))
            {
                ctx.Credentials = credentials;
                ctx.RequestTimeout = Timeout.Infinite;

                Web Web = ctx.Web;
                ctx.Load(Web, w => w.Title);
                ctx.ExecuteQueryRetry();

                System.IO.DirectoryInfo rootDir = new System.IO.DirectoryInfo(@"../Template/");

                XMLTemplateProvider provider = new XMLFileSystemTemplateProvider(rootDir.FullName + "\\pnpprovisioning", "");

                ProvisioningTemplate template = provider.GetTemplate(subsitexml);
                ProvisioningTemplateApplyingInformation ptai = new ProvisioningTemplateApplyingInformation();

                FileSystemConnector connector = new FileSystemConnector(rootDir.FullName + "\\pnpprovisioning", "");
                template.Connector = connector;
                try
                {
                    Console.WriteLine("Template is provisioning on " + WebUrl);
                    Web.ApplyProvisioningTemplate(template, ptai);


                }

                catch (Exception ex)
                {
                    String hotelSiteUrl = rootsiteUrl + "/" + hotelURL;
                    if (ex.Message.Contains("connection") || ex.Message.Contains("One or more"))
                    {
                        Console.WriteLine("Template is provisioning on " + WebUrl);
                        Web.ApplyProvisioningTemplate(template, ptai);

                        ApplyCommListToSite(WebUrl);

                    }
                    else
                    {
                        ApplyCommListToSite(WebUrl);
                        // Console.WriteLine("Error occured while appplying template." + ex.Message);
                    }
                }
            }
        }

        public static void ApplyCommListToSite(String WebUrl)
        {
            using (var ctx = new ClientContext(WebUrl))
            {
                ctx.Credentials = credentials;
                ctx.RequestTimeout = Timeout.Infinite;

                Web Web = ctx.Web;
                ctx.Load(Web, w => w.Title);
                ctx.ExecuteQueryRetry();

                System.IO.DirectoryInfo rootDir1 = new System.IO.DirectoryInfo(@"../Template/");

                XMLTemplateProvider provider1 = new XMLFileSystemTemplateProvider(rootDir1.FullName + "\\pnpprovisioning", "");

                ProvisioningTemplate template1 = provider1.GetTemplate("MyPnPProvisioningCommList.xml");
                ProvisioningTemplateApplyingInformation ptai1 = new ProvisioningTemplateApplyingInformation();

                FileSystemConnector connector1 = new FileSystemConnector(rootDir1.FullName + "\\pnpprovisioning", "");
                template1.Connector = connector1;
                try
                {
                    Console.WriteLine("Template is provisioning on " + WebUrl);
                    Web.ApplyProvisioningTemplate(template1, ptai1);

                    Console.WriteLine("Template provisioning completed successfully");

                }

                catch (Exception ex)
                {
                    String hotelSiteUrl = "";
                    if (WebUrl.ToLower().Contains("accounting") || WebUrl.ToLower().Contains("frontdesk") || WebUrl.ToLower().Contains("housekeeping") || WebUrl.ToLower().Contains("maintainence") || WebUrl.ToLower().Contains("operations") || WebUrl.ToLower().Contains("sales"))
                    {
                        hotelSiteUrl = rootsiteUrl + "/" + hotelURL + "/" + WebUrl;
                    }
                    else
                    {
                        hotelSiteUrl = rootsiteUrl + "/" + hotelURL;
                    }
                    if (ex.Message.Contains("connection") || ex.Message.Contains("One or more"))
                    {
                        ApplyCommListToSite(hotelSiteUrl);
                    }
                    else
                    {
                        //  installAppToSite(WebUrl);
                        Console.WriteLine("Error occured while appplying template." + ex.Message);
                    }
                }
            }
        }

        public static void installAppToSite(String WebUrl)
        {
            try
            {
                using (var ctx = new ClientContext(WebUrl)) 
                {
                    bool isHRSite = WebUrl.EndsWith("HR") ? true : false;
                    ctx.Credentials = credentials;

                    var appManager = new AppManager(ctx);

                    var apps = appManager.GetAvailable();
                    Console.WriteLine("Installing all the apps into the site");
                    var appinstance = apps.Where(a => a.Title == "custom-download-button-client-side-solution").FirstOrDefault();
                    if (appinstance != null)
                    {
                        //var installApp = appManager.Install(appinstance);
                        Task installTask = Task.Run(async () => await appManager.InstallAsync(appinstance));
                        installTask.Wait();
                        Console.WriteLine("App " + appinstance.Title + " installed Successfully");
                    }
                    var appinstance1 = apps.Where(a => a.Title == "headerfooter-extension-client-side-solution").FirstOrDefault();
                    if (appinstance1 != null)
                    {
                        //var installApp = appManager.Install(appinstance);
                        Task installTask = Task.Run(async () => await appManager.InstallAsync(appinstance1));
                        installTask.Wait();
                        Console.WriteLine("App " + appinstance1.Title + " installed Successfully");
                    }
                    var appinstance2 = apps.Where(a => a.Title == "HMV Solution").FirstOrDefault();
                    if (appinstance2 != null && isHRSite == false)
                    {
                        //var installApp = appManager.Install(appinstance);
                        Task installTask = Task.Run(async () => await appManager.InstallAsync(appinstance2));
                        installTask.Wait();
                        Console.WriteLine("App " + appinstance2.Title + " installed Successfully");
                    }
                    var appinstance3 = apps.Where(a => a.Title == "admin-Webparts-client-side-solution").FirstOrDefault();
                    if (appinstance3 != null)
                    {
                        //var installApp = appManager.Install(appinstance);
                        Task installTask = Task.Run(async () => await appManager.InstallAsync(appinstance3));
                        installTask.Wait();
                        Console.WriteLine("App " + appinstance3.Title + " installed Successfully");
                    }
                    var appinstance4 = apps.Where(a => a.Title == "hr-solution-client-side-solution").FirstOrDefault();
                    if (appinstance4 != null && isHRSite)
                    {
                        //var installApp = appManager.Install(appinstance);
                        Task installTask = Task.Run(async () => await appManager.InstallAsync(appinstance4));
                        installTask.Wait();
                        Console.WriteLine("App " + appinstance4.Title + " installed Successfully");
                    }
                    Console.WriteLine("All apps are installed successfully");
                }
            }
            catch (Exception ex)
            { 
                // installAppToSite(WebUrl);
                // Console.WriteLine("Error occured while appplying template." + ex.Message);
            }
        }
    
        public static void createHotel()
        {
            try
            {
                using (ClientContext ctx1 = new ClientContext(rootsiteUrl))
                {
                    ctx1.Credentials = credentials;
                    Console.WriteLine("Creating hotel");

                    WebCreationInformation wci = new WebCreationInformation();
                    wci.Url = hotelURL; // This url is relative to the url provided in the context
                    wci.Title = hotelTitle;
                    wci.Description = hotelTitle;
                    wci.UseSamePermissionsAsParentSite = true;
                    wci.WebTemplate = "STS#3";
                    wci.Language = 1033;

                    Web w = ctx1.Site.RootWeb.Webs.Add(wci);
                    ctx1.ExecuteQuery();
                    Console.WriteLine("Hotel Site Created successfully");

                    String hotelSiteUrl = rootsiteUrl + "/" + hotelURL;
                    ApplyTemplateToRootSite(hotelSiteUrl);
                    //to apply common list
                    ApplyCommListToSite(hotelSiteUrl);
                    installAppToSite(hotelSiteUrl);
                    createPages();
                    addItemsToGlobalNavigationList(hotelSiteUrl);


                    using (ClientContext ctx = new ClientContext(hotelSiteUrl))
                    {
                        ctx.Credentials = credentials;
                        //accounting site set up
                        Console.WriteLine("Creating accounting sub site");
                        WebCreationInformation accountingwci = new WebCreationInformation();
                        accountingwci.Url = "Accounting"; // This url is relative to the url provided in the context
                        accountingwci.Title = "Accounting";
                        accountingwci.Description = "Accounting";
                        accountingwci.UseSamePermissionsAsParentSite = true;
                        accountingwci.WebTemplate = "STS#3";
                        accountingwci.Language = 1033;
                        Web accountingWeb = ctx.Web.Webs.Add(accountingwci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("Accounting Site Created successfully");

                        ////to apply common list in accounting site
                        String accountingWebUrl = hotelSiteUrl + "/Accounting";
                        //String accountingWebUrl = rootsiteUrl + "/Accounting";
                        ApplyCommListToSite(accountingWebUrl);
                        installAppToSite(accountingWebUrl);
                        createsubSitePages(accountingWebUrl);
                        addItemsToGlobalNavigationList(accountingWebUrl);

                        //sales site set up
                        Console.WriteLine("Creating sales sub site");
                        WebCreationInformation saleswci = new WebCreationInformation();
                        saleswci.Url = "Sales"; // This url is relative to the url provided in the context
                        saleswci.Title = "Sales";
                        saleswci.Description = "Sales";
                        saleswci.UseSamePermissionsAsParentSite = true;
                        saleswci.WebTemplate = "STS#3";
                        saleswci.Language = 1033;
                        Web salesWeb = ctx.Web.Webs.Add(saleswci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("Sales Site Created successfully");

                        ////to apply common list in Sales site
                        String salesWebUrl = hotelSiteUrl + "/Sales";
                        //String salesWebUrl = rootsiteUrl + "/Sales";
                        ApplyCommListToSite(salesWebUrl);
                        installAppToSite(salesWebUrl);
                        createsubSitePages(salesWebUrl);
                        addItemsToGlobalNavigationList(salesWebUrl);

                        //operations site set up
                        Console.WriteLine("Creating operations sub site");
                        WebCreationInformation operationswci = new WebCreationInformation();
                        operationswci.Url = "Operations"; // This url is relative to the url provided in the context
                        operationswci.Title = "Operations";
                        operationswci.Description = "Operations";
                        operationswci.UseSamePermissionsAsParentSite = true;
                        operationswci.WebTemplate = "STS#3";
                        operationswci.Language = 1033;
                        Web operationsWeb = ctx.Web.Webs.Add(operationswci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("operations Site Created successfully");

                        ////to apply common list in Operations site
                        String operationsWebUrl = hotelSiteUrl + "/Operations";
                        //String operationsWebUrl = rootsiteUrl + "/Operations";
                        ApplyCommListToSite(operationsWebUrl);
                        installAppToSite(operationsWebUrl);
                        createsubSitePages(operationsWebUrl);
                        addItemsToGlobalNavigationList(operationsWebUrl);

                        //housekeeping site set up
                        Console.WriteLine("Creating housekeeping sub site");
                        WebCreationInformation housekeepingwci = new WebCreationInformation();
                        housekeepingwci.Url = "Housekeeping"; // This url is relative to the url provided in the context
                        housekeepingwci.Title = "Housekeeping";
                        housekeepingwci.Description = "Housekeeping";
                        housekeepingwci.UseSamePermissionsAsParentSite = true;
                        housekeepingwci.WebTemplate = "STS#3";
                        housekeepingwci.Language = 1033;
                        Web housekeepingWeb = ctx.Web.Webs.Add(housekeepingwci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("Housekeeping Site Created successfully");

                        ////to apply common list in Housekeeping site
                        String housekeepingWebUrl = hotelSiteUrl + "/Housekeeping";
                        //String housekeepingWebUrl = rootsiteUrl + "/Housekeeping";
                        ApplyTemplateToSubSite(housekeepingWebUrl, "MyPnPProvisioningHosekeepingBG.xml");
                        ApplyCommListToSite(housekeepingWebUrl);
                        installAppToSite(housekeepingWebUrl);
                        createsubSitePages(housekeepingWebUrl);
                        addItemsToGlobalNavigationList(housekeepingWebUrl);

                        //maintainence site set up
                        Console.WriteLine("Creating maintainence sub site");
                        WebCreationInformation maintainencewci = new WebCreationInformation();
                        maintainencewci.Url = "Maintainence"; // This url is relative to the url provided in the context
                        maintainencewci.Title = "Maintainence";
                        maintainencewci.Description = "Maintainence";
                        maintainencewci.UseSamePermissionsAsParentSite = true;
                        maintainencewci.WebTemplate = "STS#3";
                        maintainencewci.Language = 1033;
                        Web maintainenceWeb = ctx.Web.Webs.Add(maintainencewci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("Maintainence Site Created successfully");

                        ////to apply common list in maintainence site
                        String maintainenceWebUrl = hotelSiteUrl + "/Maintainence";
                        //String maintainenceWebUrl = rootsiteUrl + "/Maintainence";
                        ApplyTemplateToSubSite(maintainenceWebUrl, "MyPnPProvisioningMaintainenceBG.xml");
                        ApplyCommListToSite(maintainenceWebUrl);
                        installAppToSite(maintainenceWebUrl);
                        createsubSitePages(maintainenceWebUrl);
                        addItemsToGlobalNavigationList(maintainenceWebUrl);

                        //frontDesk site set up
                        Console.WriteLine("Creating FrontDesk sub site");
                        WebCreationInformation frontDeskwci = new WebCreationInformation();
                        frontDeskwci.Url = "FrontDesk"; // This url is relative to the url provided in the context
                        frontDeskwci.Title = "Front Desk";
                        frontDeskwci.Description = "Front Desk";
                        frontDeskwci.UseSamePermissionsAsParentSite = true;
                        frontDeskwci.WebTemplate = "STS#3";
                        frontDeskwci.Language = 1033;
                        Web frontDeskWeb = ctx.Web.Webs.Add(frontDeskwci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("FrontDesk Site Created successfully");

                        ////to apply common list in FrontDesk site
                        String frontDeskWebUrl = hotelSiteUrl + "/FrontDesk";
                        //String frontDeskWebUrl = rootsiteUrl + "/FrontDesk";
                        ApplyTemplateToSubSite(frontDeskWebUrl, "MyPnPProvisioningFrontdeskBG.xml");
                        ApplyCommListToSite(frontDeskWebUrl);
                        installAppToSite(frontDeskWebUrl);
                        createsubSitePages(frontDeskWebUrl);
                        addItemsToGlobalNavigationList(frontDeskWebUrl);

                        //hr site set up
                        Console.WriteLine("Creating hr sub site");
                        WebCreationInformation hrwci = new WebCreationInformation();
                        hrwci.Url = "HR"; // This url is relative to the url provided in the context
                        hrwci.Title = "HR";
                        hrwci.Description = "HR";
                        hrwci.UseSamePermissionsAsParentSite = true;
                        hrwci.WebTemplate = "STS#3";
                        hrwci.Language = 1033;
                        Web hrWeb = ctx.Web.Webs.Add(hrwci);
                        ctx.ExecuteQuery();
                        Console.WriteLine("HR Site Created successfully");

                        //to apply common list in HR site
                        String hrWebUrl = hotelSiteUrl + "/HR";
                        //String hrWebUrl = rootsiteUrl + "/HR";
                        ApplyTemplateToSubSite(hrWebUrl, "MyPnPProvisioningHRBG.xml");
                        installAppToSite(hrWebUrl);
                        createsubSitePages(hrWebUrl);
                        addItemsToGlobalNavigationList(hrWebUrl);
                    }
                        

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occured" + ex.Message);
            }
        }

        public static void createPages()
        {
            try
            {

            String hotelSiteUrl = rootsiteUrl + "/" + hotelURL;
            using (var ctx = new ClientContext(hotelSiteUrl))
            {
                ctx.Credentials = credentials;
                List Library = ctx.Web.Lists.GetByTitle("site pages");
                ctx.Load(Library);
                ctx.ExecuteQuery();

                Console.WriteLine("Creating pages");

                ClientSidePage homePage1 = ClientSidePage.Load(ctx, "Home.aspx");
                homePage1.DefaultSection.Columns.Clear();

                homePage1.AddSection(CanvasSectionTemplate.TwoColumnLeft, 1);
                homePage1.AddSection(CanvasSectionTemplate.OneColumn, 1);
                homePage1.AddSection(CanvasSectionTemplate.OneColumn, 1);
                var components = homePage1.AvailableClientSideComponents();
                var taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Recent Tasks CorpHP").FirstOrDefault();
                if (taskListWebPart != null)
                {
                    // Instantiate a client side Web part from our found Web part information
                    ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                    // Add the custom client side Web part to the page
                    homePage1.AddControl(projectsTaskListWp, homePage1.Sections[1].Columns[0]);
                }

                taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Notification CorpHP").FirstOrDefault();
                if (taskListWebPart != null)
                {
                    // Instantiate a client side Web part from our found Web part information
                    ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                    // Add the custom client side Web part to the page
                    homePage1.AddControl(projectsTaskListWp, homePage1.Sections[1].Columns[1]);
                }

                    taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Calendar CorpHP").FirstOrDefault();
                    if (taskListWebPart != null)
                    {
                        // Instantiate a client side Web part from our found Web part information
                        ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                        // Add the custom client side Web part to the page
                        homePage1.AddControl(projectsTaskListWp, homePage1.Sections[2].Columns[0]);
                    }

                    taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "TabNavCorpHP").FirstOrDefault();
                    if (taskListWebPart != null)
                    {
                        // Instantiate a client side Web part from our found Web part information
                        ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                        // Add the custom client side Web part to the page
                        homePage1.AddControl(projectsTaskListWp, homePage1.Sections[3].Columns[0]);
                    }

                    homePage1.Save();
                homePage1.Publish();

                    String[,] pagesArray = new String[5, 2] { { "Manage-Announcements", "AdminAnnouncements" }, { "Manage-Events", "AdminCalendarEvents" }, { "Manage-Employee-Spotlight", "AdminEmployeeSpotlight" }, { "Manage-QuickLinks", "AdminQuickLinks" }, { "Manage-Documents", "AdminDocuments" } };

                    for (var i = 0; i <= 4; i++)
                    {
                        var page2 = ctx.Web.AddClientSidePage(pagesArray[i, 0] + ".aspx", true);
                        ctx.ExecuteQuery();
                        page2.AddSection(CanvasSectionTemplate.OneColumn, 1);
                        ClientSidePage homePage = ClientSidePage.Load(ctx, pagesArray[i, 0] + ".aspx");
                        homePage.LayoutType = ClientSidePageLayoutType.Home;
                        var components1 = homePage.AvailableClientSideComponents();
                        var taskListWebPart1 = components1.Where(s => s.ComponentType == 1 && s.Name == pagesArray[i, 1]).FirstOrDefault();
                        if (taskListWebPart1 != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart1);
                            // Add the custom client side Web part to the page
                            homePage.AddControl(projectsTaskListWp);
                        }
                        homePage.Save();
                        homePage.Publish();
                        Console.WriteLine(pagesArray[i, 0] + "page created");
                    }
                    Console.WriteLine("Added Webparts in the homepage");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Added Webparts in the homepage");
            }
        }
        public static void createsubSitePages(String subsiteURL)
        {
            try
            {                
                using (var ctx = new ClientContext(subsiteURL))
                {
                    bool isHRSite = subsiteURL.EndsWith("HR") ? true : false;
                    ctx.Credentials = credentials;
                    List Library = ctx.Web.Lists.GetByTitle("site pages");
                    ctx.Load(Library);
                    ctx.ExecuteQuery();

                    Console.WriteLine("Creating pages");

                    ClientSidePage homePage1 = ClientSidePage.Load(ctx, "Home.aspx");
                    homePage1.DefaultSection.Columns.Clear();

                    homePage1.AddSection(CanvasSectionTemplate.TwoColumnLeft, 1);
                    homePage1.AddSection(CanvasSectionTemplate.OneColumn, 1);
                    homePage1.AddSection(CanvasSectionTemplate.OneColumn, 1);
                    var components = homePage1.AvailableClientSideComponents();
                    if (isHRSite)
                    {

                        var taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "hrEmployeeManagement").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[1].Columns[0]);
                        }

                         taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "hrNewJoinee").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[2].Columns[0]);
                        }

                         taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "HRLeaveManagement").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[3].Columns[0]);
                        }

                        var page2 = ctx.Web.AddClientSidePage("Manage-NewJoinee.aspx", true);
                        ctx.ExecuteQuery();
                        page2.AddSection(CanvasSectionTemplate.OneColumn, 1);
                        ClientSidePage homePage = ClientSidePage.Load(ctx, "Manage-NewJoinee.aspx");
                        homePage.LayoutType = ClientSidePageLayoutType.Home;
                        var components1 = homePage.AvailableClientSideComponents();
                        var taskListWebPart1 = components1.Where(s => s.ComponentType == 1 && s.Name == "").FirstOrDefault();
                        if (taskListWebPart1 != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart1);
                            // Add the custom client side Web part to the page
                            homePage.AddControl(projectsTaskListWp);
                        }
                        homePage.Save();
                        homePage.Publish();
                        Console.WriteLine("Manage-NewJoinee page created");
                    }
                    else
                    {
                        var taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Recent Tasks CorpHP").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[1].Columns[0]);
                        }

                        taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Notification CorpHP").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[1].Columns[1]);
                        }

                        taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "Calendar CorpHP").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[2].Columns[0]);
                        }

                        taskListWebPart = components.Where(s => s.ComponentType == 1 && s.Name == "TabNavCorpHP").FirstOrDefault();
                        if (taskListWebPart != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart);
                            // Add the custom client side Web part to the page
                            homePage1.AddControl(projectsTaskListWp, homePage1.Sections[3].Columns[0]);
                        }
                                       

                    homePage1.Save();
                    homePage1.Publish();

                    String[,] pagesArray = new String[5, 2] { { "Manage-Announcements", "AdminAnnouncements" }, { "Manage-Events", "AdminCalendarEvents" }, { "Manage-Employee-Spotlight", "AdminEmployeeSpotlight" }, { "Manage-QuickLinks", "AdminQuickLinks" }, { "Manage-Documents", "AdminDocuments" } };

                    for (var i = 0; i <= 4; i++)
                    {
                        var page2 = ctx.Web.AddClientSidePage(pagesArray[i, 0] + ".aspx", true);
                        ctx.ExecuteQuery();
                        page2.AddSection(CanvasSectionTemplate.OneColumn, 1);
                        ClientSidePage homePage = ClientSidePage.Load(ctx, pagesArray[i, 0] + ".aspx");
                        homePage.LayoutType = ClientSidePageLayoutType.Home;
                        var components1 = homePage.AvailableClientSideComponents();
                        var taskListWebPart1 = components1.Where(s => s.ComponentType == 1 && s.Name == pagesArray[i, 1]).FirstOrDefault();
                        if (taskListWebPart1 != null)
                        {
                            // Instantiate a client side Web part from our found Web part information
                            ClientSideWebPart projectsTaskListWp = new ClientSideWebPart(taskListWebPart1);
                            // Add the custom client side Web part to the page
                            homePage.AddControl(projectsTaskListWp);
                        }
                        homePage.Save();
                        homePage.Publish();
                        Console.WriteLine(pagesArray[i, 0] + "page created");
                    }
                }
                    Console.WriteLine("Added Webparts in the homepage");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Added Webparts in the homepage");
            }
        }

        public static void addItemsToGlobalNavigationList(String siteURL)
        {

            using (ClientContext clientContext = new ClientContext(siteURL))
            {
                clientContext.Credentials = credentials;
                List oList = clientContext.Web.Lists.GetByTitle("GlobalNavigationList");

                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

                ListItem oListItem = oList.AddItem(itemCreateInfo);
                oListItem["Title"] = "Home";
                oListItem["MegaMenuCategory"] = "Home";
                oListItem["MegaMenuItemName"] = "Home";
                oListItem["MegaMenuItemUrl"] = siteURL;
                oListItem.Update();

                ListItem oListItem1 = oList.AddItem(itemCreateInfo);
                oListItem1["Title"] = "Corporate";
                oListItem1["MegaMenuCategory"] = "Corporate";
                oListItem1["MegaMenuItemName"] = "Corporate";
                oListItem1["MegaMenuItemUrl"] = rootsiteUrl;
                oListItem1.Update();


                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem2 = oList.AddItem(itemCreateInfo);
                    oListItem2["Title"] = "HR";
                    oListItem2["MegaMenuCategory"] = "HR";
                    oListItem2["MegaMenuItemName"] = "HR";
                    oListItem2["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/HR";
                    oListItem2.Update();
                }

                if (siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem2 = oList.AddItem(itemCreateInfo);
                    oListItem2["Title"] = "Hotels";
                    oListItem2["MegaMenuCategory"] = "Hotels";
                    oListItem2["MegaMenuItemName"] = "Hotels";
                    oListItem2["MegaMenuItemUrl"] = "#";
                    oListItem2.Update();

                    ListItem oListItem10 = oList.AddItem(itemCreateInfo);
                    oListItem10["Title"] = "Hotels";
                    oListItem10["MegaMenuCategory"] = "Hotels";
                    oListItem10["MegaMenuItemName"] = hotelTitle;
                    oListItem10["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL;
                    oListItem10.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem3 = oList.AddItem(itemCreateInfo);
                    oListItem3["Title"] = "No";
                    oListItem3["MegaMenuCategory"] = "Departments";
                    oListItem3["MegaMenuItemName"] = "Departments";
                    oListItem3["MegaMenuItemUrl"] = "#";
                    oListItem3.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem4 = oList.AddItem(itemCreateInfo);
                    oListItem4["Title"] = "No";
                    oListItem4["MegaMenuCategory"] = "Departments";
                    oListItem4["MegaMenuItemName"] = "Accounting";
                    oListItem4["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/Accounting";
                    oListItem4.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem5 = oList.AddItem(itemCreateInfo);
                    oListItem5["Title"] = "No";
                    oListItem5["MegaMenuCategory"] = "Departments";
                    oListItem5["MegaMenuItemName"] = "Front Desk";
                    oListItem5["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/FrontDesk";
                    oListItem5.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem6 = oList.AddItem(itemCreateInfo);
                    oListItem6["Title"] = "No";
                    oListItem6["MegaMenuCategory"] = "Departments";
                    oListItem6["MegaMenuItemName"] = "Housekeeping";
                    oListItem6["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/Housekeeping";
                    oListItem6.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("operations") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem7 = oList.AddItem(itemCreateInfo);
                    oListItem7["Title"] = "No";
                    oListItem7["MegaMenuCategory"] = "Departments";
                    oListItem7["MegaMenuItemName"] = "Maintainence";
                    oListItem7["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/Maintainence";
                    oListItem7.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("sales"))
                {
                    ListItem oListItem8 = oList.AddItem(itemCreateInfo);
                    oListItem8["Title"] = "No";
                    oListItem8["MegaMenuCategory"] = "Departments";
                    oListItem8["MegaMenuItemName"] = "Operations";
                    oListItem8["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/Operations";
                    oListItem8.Update();
                }

                if (siteURL.Contains(hotelURL) || siteURL.ToLower().Contains("accounting") || siteURL.ToLower().Contains("frontdesk") || siteURL.ToLower().Contains("housekeeping") || siteURL.ToLower().Contains("maintainence") || siteURL.ToLower().Contains("operations"))
                {
                    ListItem oListItem9 = oList.AddItem(itemCreateInfo);
                    oListItem9["Title"] = "No";
                    oListItem9["MegaMenuCategory"] = "Departments";
                    oListItem9["MegaMenuItemName"] = "Sales";
                    oListItem9["MegaMenuItemUrl"] = rootsiteUrl + "/" + hotelURL + "/Sales";
                    oListItem9.Update();
                }

                clientContext.ExecuteQuery();
            }
        }

        public static String RemoveSpecialChar(String str)
        {
            str = Convert.ToString(str);

            if (str.Trim() == "")
                return str;

            str = str.Replace("&", "&amp;");
            str = str.Replace(">", "&gt;");
            str = str.Replace("<", "&lt;");
            str = str.Replace("\"", "&quot;");
            str = str.Replace("'", "&#039;");
            return str;
        }

        #region Common Helpers
        private static SharePointOnlineCredentials GetCredentials(String userName, String password)
        {
            //String userName = System.Configuration.ConfigurationManager.AppSettings["UserName"];
            //String password = System.Configuration.ConfigurationManager.AppSettings["UserPassword"];

            var securePassword = new System.Security.SecureString();
            foreach (var c in password)
            {
                securePassword.AppendChar(c);
            }

            return new SharePointOnlineCredentials(userName, securePassword);
        }
        #endregion

        #region if (String.IsNullOrWhiteSpace(O365userName))
            //{
            //    Console.WriteLine("Please provide correct UserName.");
            //}
            //else
            //{
            //    //get password
            //    Console.WriteLine("Please provide Password");
            //    String pass = "";
            //    do
            //    {
            //        ConsoleKeyInfo key = Console.ReadKey(true);
            //        // Backspace Should Not Work
            //        if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
            //        {
            //            pass += key.KeyChar;
            //            Console.Write("*");
            //        }
            //        else
            //        {
            //            if (key.Key == ConsoleKey.Backspace && pass.Length > 0)
            //            {
            //                pass = pass.SubString(0, (pass.Length - 1));
            //                Console.Write("\b \b");
            //            }
            //            else if (key.Key == ConsoleKey.Enter)
            //            {
            //                break;
            //            }
            //        }
            //    } while (true);
            //    if (String.IsNullOrWhiteSpace(O365password))
            //    {
            //        Console.WriteLine("Please provide correct Password.");
            //    }
            //    else
            //    {
                  

            //    }

            //}
    #endregion
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Xml;
using System.IO;
using System.ServiceModel.Syndication;
using System.Xml.Linq;

namespace OutlookRulesExport
{
    class Program
    {
        static void Main(string[] args)
        {
            Store store;
            String storeName;

            if (args.Count() == 0)
            {
                PrintHelp();
                return;
            }

            // print the stores(mailboxes) available
            if (args[0] == "-s")
            {
                GetStores();
                return;
            }

            // print out rules in CSV format
            if (args[0] == "-c")
            {
                if (args.Count() >= 2)
                {
                    storeName = args[1];

                    store = GetStore(storeName);
                    if (store != null)
                    {
                        List<MyRule> rulesList = GetRules(store);
                        PrintCSV(rulesList);
                    }
                }

                return;
            }

            // default, print rules to xml file
            storeName = args[0];
            store = GetStore(storeName);
            if (store != null)
            {
                List<MyRule> rulesList = GetRules(store);
                CleanRuleActions(rulesList, storeName);

                string fileName = "rules.xml";
                PrintXML(rulesList, fileName);
            }
        }

        private static void PrintHelp()
        {
            Console.WriteLine(String.Empty);
            Console.WriteLine("Exports rules from oulook so they can be directly imported into Gmail as filters to a rules.xml file in the current directory. No editing required.");
            Console.WriteLine(String.Empty);
            Console.WriteLine("Usage: OutlookRulesExport.exe [options] [mailbox]");
            Console.WriteLine(String.Empty);
            Console.WriteLine("Example: OutlookRulesExport.exe example@example.com");
            Console.WriteLine("Example: OutlookRulesExport.exe -c example@example.com");
            Console.WriteLine("Example: OutlookRulesExport.exe -s");
            Console.WriteLine(String.Empty);
            Console.WriteLine("-s: returns a list of stores (mailboexes) available");
            Console.WriteLine("-c: prints rules to the console in CSV format (unedited)");
        }

        /// <summary>
        /// Returns a list of stores (mailboxes and pst files) available in outlook
        /// Some or all of these stores can contain rules that can be exported
        /// </summary>
        public static void GetStores()
        {
            Microsoft.Office.Interop.Outlook.Application app = new Application();
            Stores stores = app.Session.Stores;

            foreach (Store s in stores)
            {
                Console.WriteLine(s.DisplayName);
            }
        }

        /// <summary>
        /// Gets the Outlook store (mailbox) for the given string if one exists.
        /// </summary>
        /// <param name="storeName"></param>
        /// <returns>Null if no store is found</returns>
        public static Store GetStore(string storeName)
        {
            Microsoft.Office.Interop.Outlook.Application app = new Application();
            Stores stores = app.Session.Stores;

            if (stores.Count > 0)
            {
                try
                {
                    Store s = stores[storeName];

                    if (s != null)
                    {
                        return s;
                    }
                    else
                    {
                        Console.WriteLine("Invalid mailbox");
                        PrintHelp();
                    }
                }
                catch (System.Runtime.InteropServices.COMException e)
                {
                    Console.WriteLine("No such mailbox");
                    PrintHelp();
                }
            }
            else
            {
                Console.WriteLine("No mailboxes in Outlook");
                PrintHelp();
            }

            return null;
        }

        /// <summary>
        /// Gets a list of rules and associated actions from the given store
        /// Currently only working for rules defined on from addresses, will add support for other rules soon
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static List<MyRule> GetRules(Store s)
        {
            Rules rules = s.GetRules();
            List<MyRule> rulesList = new List<MyRule>();

            foreach (Rule r in rules)
            {
                string address = "";
                string path ="";

                // condition from email address & move to folder
                if (r.Conditions.From.Recipients.Count > 0)
                {
                    // voodo to extract smtp email address from outlook
                    try
                    {
                        OlAddressEntryUserType addressType = r.Conditions.From.Recipients[1].AddressEntry.AddressEntryUserType;

                        if ((addressType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) || (addressType == OlAddressEntryUserType.olExchangeUserAddressEntry))
                        {
                            address = r.Conditions.From.Recipients[1].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else
                        {
                            if (addressType == OlAddressEntryUserType.olSmtpAddressEntry)
                            {
                                address = r.Conditions.From.Recipients[1].AddressEntry.Address;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    // capture the action
                    // TODO check if the rule is in a error state
                    try
                    {
                        if (r.Actions[1].ActionType == OlRuleActionType.olRuleActionMoveToFolder)
                        {
                            if (r.Actions.MoveToFolder.Enabled)
                            {
                                MAPIFolder folder = r.Actions.MoveToFolder.Folder;
                                if (folder != null)
                                {
                                    path = folder.FolderPath;

                                    MyRule myRule = new MyRule();
                                    myRule.Address = address;
                                    myRule.Action = path;

                                    rulesList.Add(myRule);
                                }
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }
                }
            }

            return rulesList;
        }

        /// <summary>
        /// Print the rules in CSV format to the console
        /// </summary>
        /// <param name="rules"></param>
        public static void PrintCSV(List<MyRule> rules)
        {
            // print the rules
            foreach (MyRule mr in rules)
            {
                Console.WriteLine(mr.Address + "," + mr.Action);
            }
        }

        /// <summary>
        /// Cleans the actions in the rule list to remove the storename and formats for 
        /// Google.
        /// </summary>
        /// <param name="rules"></param>
        /// <param name="storeName"></param>
        public static void CleanRuleActions(List<MyRule> rules, String storeName)
        {
            foreach (MyRule r in rules)
            {
                // remove the store name from the action path
                r.Action = r.Action.Replace("\\\\"+storeName+"\\", String.Empty);

                // swap backslash for forwardslash to play nice with google
                r.Action = r.Action.Replace("\\", "/");
            }
        }

        /// <summary>
        /// Prints the rules to an XML file called rules.xml in the current dir. 
        /// This was a real pain to get working correctly.
        /// </summary>
        /// <param name="rules"></param>
        public static void PrintXML(List<MyRule> rules, string filename)
        {
            SyndicationFeed feed = new SyndicationFeed();
            XmlQualifiedName key = new XmlQualifiedName("apps", "http://www.w3.org/2000/xmlns/");
            String googleNs = "http://schemas.google.com/apps/2006";
            feed.AttributeExtensions.Add(key, googleNs);

            List<SyndicationItem> feedItems = new List<SyndicationItem>();

            foreach (MyRule r in rules)
            {
                SyndicationItem atom = new SyndicationItem();
                atom.Title = new TextSyndicationContent(r.Address, TextSyndicationContentKind.Plaintext);
                atom.Categories.Add(new SyndicationCategory("filter"));
                atom.Content = new TextSyndicationContent(String.Empty);

                //atom.ElementExtensions.Add(new SyndicationElementExtension("property", googleNs, ""));

                XName n = XName.Get("property", googleNs);
                XElement el1a = new XElement(n);
                XAttribute at1a = new XAttribute("name", "from");
                el1a.Add(at1a);
                XAttribute at1b = new XAttribute("value", r.Address);
                el1a.Add(at1b);
                atom.ElementExtensions.Add(el1a);

                XElement el2a = new XElement(n);
                XAttribute at2a = new XAttribute("name", "label");
                el2a.Add(at2a);
                XAttribute at2b = new XAttribute("value", r.Action);
                el2a.Add(at2b);
                atom.ElementExtensions.Add(el2a);

                XElement el3a = new XElement(n);
                XAttribute at3a = new XAttribute("name", "shouldArchive");
                el3a.Add(at3a);
                XAttribute at3b = new XAttribute("value", "true");
                el3a.Add(at3b);
                atom.ElementExtensions.Add(el3a);

                feedItems.Add(atom);
            }

            feed.Items = feedItems;

            StreamWriter sw = new StreamWriter(filename);

            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Indent = true;
            settings.IndentChars = "\t";

            XmlWriter xw = XmlWriter.Create(sw, settings);

            Atom10FeedFormatter atomFormatter = new Atom10FeedFormatter(feed);            
            atomFormatter.WriteTo(xw);

            xw.Flush();
            xw.Close();
        }
    }

    class MyRule
    {
        string address;

        public string Address
        {
            get { return address; }
            set { address = value; }
        }
        string action;

        public string Action
        {
            get { return action; }
            set { action = value; }
        }
    }
}

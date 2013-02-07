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
                        List<MyRule> rulesList = GetRules(store, storeName);
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
                List<MyRule> rulesList = GetRules(store, storeName);

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
        public static List<MyRule> GetRules(Store s, string storeName)
        {
            Rules rules = s.GetRules();
            List<MyRule> rulesList = new List<MyRule>();

            foreach (Rule r in rules)
            {
                if (r.Enabled)
                {
                    MyRule mr = new MyRule();

                    ParseFromAddresses(r, mr);
                    ParseLabelMove(r, mr, storeName);
                    ParseLabelCopy(r, mr, storeName);
                    ParseSubject(r, mr);
                    ParseBody(r, mr);

                    rulesList.Add(mr);
                }
            }

            return rulesList;
        }

        private static void ParseBody(Rule r, MyRule mr)
        {
            if (r.Conditions.Body.Enabled)
            {
                string[] temp = r.Conditions.Body.Text;

                for (int i = 0; i < temp.Length; i++)
                {
                    mr.HasTheWord += temp[i];

                    if (i != temp.Length - 1)
                        mr.HasTheWord += " OR ";
                }
            }
        }

        private static void ParseSubject(Rule r, MyRule mr)
        {
            if (r.Conditions.Subject.Enabled)
            {
                string[] temp = r.Conditions.Subject.Text;

                for (int i = 0; i < temp.Length; i++)
                {
                    mr.Subject += temp[i];

                    if (i != temp.Length - 1)
                        mr.Subject += " OR ";
                }
            }
        }

        /// <summary>
        /// Parses a rule for MoveToFolder actions
        /// Converts this to a GMail label also sets the ShouldArchive option on the 
        /// gmail filter which applies the label makes the mail skip the gmail inbox
        /// 
        /// </summary>
        /// <param name="r"></param>
        /// <param name="mr"></param>
        private static void ParseLabelMove(Rule r, MyRule mr, string storeName)
        {
            try
            {
                if (r.Actions[1].ActionType == OlRuleActionType.olRuleActionMoveToFolder)
                {
                    if (r.Actions.MoveToFolder.Enabled)
                    {
                        MAPIFolder folder = r.Actions.MoveToFolder.Folder;
                        if (folder != null)
                        {
                            mr.Label = CleanRuleActions(folder.FolderPath, storeName);
                            mr.ShouldArchive = true;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        /// <summary>
        /// Parses a rule for the CopyToFolder action
        /// Converts the action of a GMail label but does not set the ShouldArchive option
        /// </summary>
        /// <param name="r"></param>
        /// <param name="mr"></param>
        private static void ParseLabelCopy(Rule r, MyRule mr, string storeName)
        {
            try
            {
                if (r.Actions.CopyToFolder.Enabled)
                {
                    MAPIFolder folder = r.Actions.CopyToFolder.Folder;
                    if (folder != null)
                    {
                        mr.Label = CleanRuleActions(folder.FolderPath, storeName);
                    }
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        /// <summary>
        /// condition from email address & move to folder
        /// mupports multiple addresses in the from
        /// </summary>
        /// <param name="r"></param>
        /// <param name="mr"></param>
        private static void ParseFromAddresses(Rule r, MyRule mr)
        {
            if (r.Conditions.From.Recipients.Count > 0)
            {
                for (int i = 1; i <= r.Conditions.From.Recipients.Count; i++)
                {
                    string temp = "";
                    // voodo to extract email addresses
                    try
                    {
                        OlAddressEntryUserType addressType = r.Conditions.From.Recipients[i].AddressEntry.AddressEntryUserType;

                        if ((addressType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry) || (addressType == OlAddressEntryUserType.olExchangeUserAddressEntry))
                        {
                            temp = r.Conditions.From.Recipients[i].AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        }
                        else
                        {
                            if (addressType == OlAddressEntryUserType.olSmtpAddressEntry)
                            {
                                temp = r.Conditions.From.Recipients[i].AddressEntry.Address;
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Console.WriteLine(ex);
                    }

                    // compose the address string if there are mutlitple addresses in the from
                    if (!String.IsNullOrEmpty(temp))
                    {
                        if (i == 1)
                        {
                            mr.FromAddress += temp;
                        }
                        else
                        {
                            mr.FromAddress += "," + temp;
                        }
                    }
                }
            }
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
                Console.WriteLine(mr.FromAddress + ";" + mr.Label);
            }
        }

        /// <summary>
        /// Cleans the actions in the rule list to remove the storename and formats for 
        /// Google.
        /// </summary>
        /// <param name="rules"></param>
        /// <param name="storeName"></param>
        public static string CleanRuleActions(string path, String storeName)
        {
            // remove the store name from the action path
            path =  path.Replace("\\\\"+storeName+"\\", String.Empty);

            // swap backslash for forwardslash to play nice with google
            path = path.Replace("\\", "/");

            return path;
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
                atom.Title = new TextSyndicationContent(r.FromAddress, TextSyndicationContentKind.Plaintext);
                atom.Categories.Add(new SyndicationCategory("filter"));
                atom.Content = new TextSyndicationContent(String.Empty);

                //atom.ElementExtensions.Add(new SyndicationElementExtension("property", googleNs, ""));
                XName n = XName.Get("property", googleNs);

                // conditions
                bool conditionSet = false;
                if (!String.IsNullOrEmpty(r.FromAddress))
                {
                    XElement el1a = new XElement(n);
                    XAttribute at1a = new XAttribute("name", "from");
                    el1a.Add(at1a);
                    XAttribute at1b = new XAttribute("value", r.FromAddress);
                    el1a.Add(at1b);
                    atom.ElementExtensions.Add(el1a);
                    
                    conditionSet = true;
                }

                if (!String.IsNullOrEmpty(r.Subject))
                {
                    XElement el4a = new XElement(n);
                    XAttribute at4a = new XAttribute("name", "subject");
                    el4a.Add(at4a);
                    XAttribute at4b = new XAttribute("value", r.Subject);
                    el4a.Add(at4b);
                    atom.ElementExtensions.Add(el4a);
                    
                    conditionSet = true;
                }

                if (!String.IsNullOrEmpty(r.HasTheWord))
                {
                    XElement el4a = new XElement(n);
                    XAttribute at4a = new XAttribute("name", "hasTheWord");
                    el4a.Add(at4a);
                    XAttribute at4b = new XAttribute("value", r.HasTheWord);
                    el4a.Add(at4b);
                    atom.ElementExtensions.Add(el4a);

                    conditionSet = true;
                }

                // actions - only apply if a condition has been set
                if (conditionSet)
                {
                    if (!String.IsNullOrEmpty(r.Label))
                    {
                        XElement el2a = new XElement(n);
                        XAttribute at2a = new XAttribute("name", "label");
                        el2a.Add(at2a);
                        XAttribute at2b = new XAttribute("value", r.Label);
                        el2a.Add(at2b);
                        atom.ElementExtensions.Add(el2a);
                    }

                    if (r.ShouldArchive)
                    {
                        XElement el3a = new XElement(n);
                        XAttribute at3a = new XAttribute("name", "shouldArchive");
                        el3a.Add(at3a);
                        XAttribute at3b = new XAttribute("value", r.ShouldArchive.ToString());
                        el3a.Add(at3b);
                        atom.ElementExtensions.Add(el3a);
                    }


                    feedItems.Add(atom);
                }
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
        string toAddress;
        string subject;

        public string Subject
        {
            get { return subject; }
            set { subject = value; }
        }

        string hasTheWord;

        public string HasTheWord
        {
            get { return hasTheWord; }
            set { hasTheWord = value; }
        }
        string doesNotHaveTheWord;
        string hasAttachment;
        string shouldMarkAsRead;
        bool shouldArchive;

        public bool ShouldArchive
        {
            get { return shouldArchive; }
            set { shouldArchive = value; }
        }
        string shouldStar;
        string shouldTrash;
        string shouldAlwaysMarkAsImportant;

        string fromAddress;
        public string FromAddress
        {
            get { return fromAddress; }
            set { fromAddress = value; }
        }

        string label;
        public string Label
        {
            get { return label; }
            set { label = value; }
        }
    }
}

##Outlook Rules Export To Gmail##
* Simple tool to export rules from Outlook and import them into Google Gmail as filters
* Tested on Win7 64bit, Win8 64bit & Outlook 2007, 2010 and 2013
* It also outputs rules in CSV to the command line so you can script it if you want
* http://bcleary.github.com/outlook-gmail-rules-export/

##Download##
* [0.0.1 - Stable version - only supports from move-to-folder rules](http://bit.ly/W5OAZi)
* [0.0.2 - Beta version - supports all rules listed below](http://bit.ly/11WLoo1)

##System Requirements##
* Office 2007 or later
* .Net 4.5 http://www.microsoft.com/en-ca/download/details.aspx?id=30653

##Usage##
* Run the tool from the command line;
     * OutlookRulesExport.exe example@example.com
     * This will create a `rules.xml` file in the current directory containing the rules
     * You might get a warning from outlook asking you to allow access, if so click ok.
* Log into Gmail
* Enable gmail filters import in Gmail
      * Setting -> Labs enable "Filter import/export" and click save changes
* Backup your existing Gmail filters
     * Settings -> Filters -> "Select All" -> "Export"
* Import the rules as Gmail filters
      * Settings -> Filters click "Import Filters" and import `rules.xml`
      * Importing existing filters into gmail is handled gracefully, no duplicate filters should be created, but its always a good idea to backup your filters

##Supported Rule Types##
* The tool currently only works for the following types of rules (will add support for other types of rules as i need them or as requested)
    * Condition: "From Address" Actions: "Move-To-Folder | Copy-To-Folder"
    * Condition: "Subject Contains" Actions: "Move-To-Folder | Copy-To-Folder"
    * Condition: "Body Contains" Actions: "Move-To-Folder | Copy-To-Folder"

##Credits##
* This is a rewrite of the https://github.com/iloveitaly/outlook-gmail-rules-migration project from ruby and vb to c#
* Thanks to iloveitaly (Michael Bianco) https://github.com/iloveitaly for his work figuring out the various ways Outlook stores its rules https://github.com/iloveitaly/outlook-gmail-rules-migration 

##Outlook Rules Export To Gmail##
* Simple tool to export rules from Outlook and import them into Google Gmail as filters
* Tested on Win7 64bit, Win8 64bit & Outlook 2013, 2010 but it should work on 2007
* It also outputs rules in CSV to the command line so you can script it if you want

##Download##
* [0.0.1 - Stable version](https://github.com/bcleary/outlook-gmail-rules-export/blob/gh-pages/downloads/0.0.1/OutlookRulesExport.exe?raw=true)
* [0.0.2 - Beta version](https://github.com/bcleary/outlook-gmail-rules-export/blob/gh-pages/downloads/0.0.2/OutlookRulesExport.exe?raw=true)

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

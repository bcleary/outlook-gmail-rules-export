outlook-gmail-rules-export
==========================

* This is a tool to export email rules from Outlook into a file format that can be imported to Gmail. 
* The tool currently only works for "From Address - Move to Folder" rules, will add support for other types of rules as i need them.

Usage
==========================

* Download the tool from https://github.com/downloads/bcleary/outlook-gmail-rules-export/OutlookRulesExport.exe
* Run the tool from the command line;
  * OutlookRulesExport.exe example@example.com
  * This will create a `rules.xml` file in the current directory containing the rules
* Log into Gmail
* Enable gmail filters import in Gmail
  * Setting -> Labs enable "Filter import/export" and click save changes
* Import the rules as Gmail filters 
  * Settings -> Filters click "Import Filters" and import `rules.xml`
  * Importing existing filters into gmail is handled gracefully, no duplicate filters should be created but its always a good idead to backup your filters
##outlook-gmail-rules-export##

* This is a command line tool to export email rules from Outlook into a file format that can be imported to Gmail. 
* The tool currently only works for "From Address - Move to Folder" type rules, will add support for other types of rules as i need them.
* I have tested this on Win 7 64bit & Outlook 2010 but it should work on 2007 also
* It also outputs rules in CSV to the command line so you can script it if you want
* Please report any bugs

##Usage##

* Build it or download the binary from https://github.com/downloads/bcleary/outlook-gmail-rules-export/OutlookRulesExport.exe
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

##Credits##
* This is a rewrite of the https://github.com/iloveitaly/outlook-gmail-rules-migration project from ruby and vb to c#
* Thanks to iloveitaly (Michael Bianco) https://github.com/iloveitaly for his work figuring out the various ways Outlook stores its rules https://github.com/iloveitaly/outlook-gmail-rules-migration 
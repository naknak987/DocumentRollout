# DocumentRollout
A windows tool for copying, moving, renaming and deleting files on a structured file share.

### Usage
On a corporate shared drive. Say you have multiple facilities or department that each have their own folder on the shared drive, and you need to quickly push out a report or form (some kind of document) to each facility. You point the application at the document you want to "rollout" and point the application to the location inside one of the facility folders. Then check the boxes next to each facility you want to roll it out to and push the rollout button. Within seconds the document is copied out to each checked location. 

### Setup
Before you can use the Document Rollout application, you need to modify the settings.ini file. You can modify it with any text editor. Notepad works well. 
The structure of the settings.ini file is very important to the correct operation of the application and the first 7 lines should not be modified, nor should there be any extra lines added above line 8. 

#### Line 7 Folder Listing
After line 7 there are 3 lines that look like the fallowing
```text
^facList1^
^facList2^
^facList3^
```
This allows you to organize the facilities or departments on your shared drive into different lists. The list names can be changed but the carrots `^` must be before and after the list name. 
Under each list name, you can add the list items. Let’s say we have a mapped shared drive `X:\` that is used for multiple facilities or departments in your organization. In `X:\` you have a folder for each facility or department. So now we have...
```text
X:\facility_1
x:\facility_2
x:\facility_3
... and so on
```
You could change your settings.ini file so that the facility listing looks like this.
```text
^Main Facilities^
facility_1
facility_2
^Secondary Facilities^
facility_3
facility_4
^Tertiary Facilities^
facility_5
facility_6
```
Note that the facility names in the settings.ini must match the folder name for that facility. This is how the program keeps track of what it’s doing.

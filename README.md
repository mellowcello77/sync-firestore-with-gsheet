## This apps script was created to import and export between google sheets and Firestore

## Original website where this was found:
https://liftcodeplay.com/2020/02/08/exporting-and-importing-data-to-from-firestore-using-google-sheets/comment-page-1/?unapproved=17341&moderation-hash=cf574195653a2711a771549ce26ccd9e#comment-17341

## The script is a javascript google apps script be used in a gsheet script. 
It connects to Firease, collects the documents from a collection and writes them into your google sheet. 
The script can be copied and pasted as is, just the credentials to Firebase/Firestore needs adding. 
Blog for more information: https://github.com/grahamearley/FirestoreGoogleAppsScript

## Required variables to access firestore is in the script:
const key = '-----BEGIN PRIVATE KEY-----\YourPrivateKeyHere\n-----END PRIVATE KEY-----\n' // TODO - enter your private key here

const email = 'YourGoogleServiceAccountEmail@gserviceaccount.com' // TODO - enter your email here

const projectId = 'YourProjectNameID' // TODO - Enter your project ID here

## Opened a Sheet
Clicked Tools > Script Editor
Typed up that code. In your case, just paste it!
Followed the quick start mentioned here, to install the library, get my secrets and add the secrets to my script https://github.com/grahamearley/FirestoreGoogleAppsScript
You’ll see several comments with TODO that you need to fill out
Window popped up occasionally asking for permission to connect to my GCP / Firestore instance. I gave it permission
Made sure to clicked Save
How it works
There’s two parts to my script – exports and imports.

Exports:
This is largely based on the content of the video but I took it a bit further. Key things I added were:

Column A is the ID of the doc.
If it’s blank the document will be added and that field will show the ID from Firestore
If it’s not blank it will update the document in Firestore
Column H (the far right) I called toDelete. If that field is non-empty it will delete the row in Firestore, then the sheet
The video just grabbed 100 rows. My sheet grabs the actual data in the sheet

Imports:
I added another menu item and added functionality to:

Import records that don’t exist in the sheet to the sheet
Update existing records in the sheet. The import will match up the properties to the correct column
Note that it does not handle the situation where the record is not in Firestore but is in the sheet

All the code can be found here, in this repo:
https://github.com/bcnzer/firebase-googlesheets-importexport

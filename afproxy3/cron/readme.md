## 

This folder is hosted on a seperate Val.run instance.

it is grouped inside this directory to make researching the documents and code structure easier. do not use the outdated ~/afproxy folder found in root.

this cron is the v0 SQL implementation for the work orders page. 

our objective is to extract the work order uuid from AppFolios v0 api (see .api folder) in order to correctly resolve attachments, photos, and notes on the work order and billing modals using our own indexing - association schematic found in the sql. 
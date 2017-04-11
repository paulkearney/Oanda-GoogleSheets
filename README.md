# Oanda-GoogleSheets
A Google Sheet that imports data from Oanda's v20 API for historical analysis

## Purpose
I created this Google Sheet for a few reasons:

1. Allow me to quickly review my open trades without needing to remote in to my VPS.
2. Allow me to perform some historical analysis on my trade history.
3. I like archiving all of my data somewhere.

## Overview
This sheet will, every 5 minutes while it is open, connect to Oanda and update three worksheets of data:

1. `Account` worksheet: Account overview
2. `Open Trades` worksheet: Lists all open trades in the account
3. `History` worksheeet: A list of all closed trades in the account.
To view the code (and make sure I'm not secretly stealing your credentials or your account), you can view the file `OandaHistory.js` in this repository. You can also, once you've completed the steps below, select the `Tools` menu, then `Script Editor`.

## Getting Started
In order to use this, you will need:

1. To be an Oanda customer.
2. To have an Oanda v20 account. This will not work with their older v1 accounts.
3. An Oanda API Key. To obtain one, head over to https://www.oanda.com/account/tpa/personal_token. Click generate and make note of your API key. **You will not be able to retrieve the API key again once you leave this page. Make a note of the API key and store it in a safe, secure location. Anyone who has access to this key has access to your account. Do not share the key.**
4. A Google Docs account. If you don't have one, sign up for one at https://docs.google.com.

Next, you will need to copy the Google Sheet into your own Google account.

1. Go to this document: https://docs.google.com/spreadsheets/d/1u39rB8u0aD__9O_aJbHU2wtIgy83h36EXpIZiV_bIAg/
2. Make a copy into your own account. Select the `File` menu, then Make a Copy. Give the document a name.
3. Enter your Oanda API Key on the sheet. On the `Account` tab, scroll down to row 99. In the second column, enter the API key you received from Oanda. *This is in row 99 so it is not visible unless you explicitly scroll down. Once set you do not need to change this.*
4. Enter your Oanda v20 account number on the sheet. On the `Account` tab, scroll down to row 99. In the second column, enter your v20 account number. You can find it by going to https://www.oanda.com/funding/, selecting the account in the list on the left, then copying the `v20 Account Number`. Note this is not your MT4 account number.
5. Click the `Oanda` menu, then select `Update`. 
6. You will be prompted to give 'Oanda History' permission to run. Click 'Continue'.
7. You will see the `Open Trades` and `History` sheets get updated. The first load of history can take a while depending on how much history you have. 

## Updating the Data
The data automatically refreshes itself every 5 minutes. You can also trigger a manual refresh by clicking the `Oanda` menu and selecting `Update`.

## Troubleshooting
This sheet has been tested with a few live and demo accounts, but it is possible I didn't account for every case. If you see something strange, it is most likely related to the `History` sheet. If it seems like this data isn't refreshing correctly, you can select all of the data on the `History` sheet by pressing `CTRL+A` (`CMD+A` on a Mac). Then right click and select `Delete Rows X-Y` (`X-Y` will be numbers). This will force the `History` sheet to be repopulated on the next update.

## Issues? Questions?
If you have any issues or questions, you can get me on TFL365 Slack @pk, or leave an issue on GitHub.

## Future Updates
Plans for the future:

1. I've modified my robots (Finch) so that, when taking a second trade, it includes the first trade's order number in the comments. I plan to add another sheet that ties the first and second trades together so you can analyze them as a set.
2. Add support for multiple accounts in one spreadsheet. For now, if you have multiple v20 accounts, you need to create multiple spreadsheets.

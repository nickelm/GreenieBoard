# GreenieBoard for DCS

GreenieBoard is a Python scraper that takes output from LSO BOT to log carrier landing stats and generate a greenie board on Google Sheets. The script will read data from a CSV file that LSO BOT will generate. This data will be used to calculate lifetime and per month landings and point statistics for all pilots. Finally, the script will check the landing events against the existing ones in the Google Sheets and will add any new ones. If there are new ones, the greenie board for that month will be updated. 

Below is the overall process for installing and running GreenieBoard:
1. Modify LSO Bot
2. Install GreenieBoard dependencies
3. Setting up the Google Service account
4. Creating the feed and board sheets
5. Configuring GreenieBoard
6. Running and scheduling GreenieBoard

Read on for details on each of these points.

## Modifying LSO Bot

The original LSO BOT script (https://github.com/YoloWingPixie/lsobot) only has Discord integration and a basic logfile output. To produce structured CSV output, you need to make the following modifications to the LSO BOT files:

Add the following to line 58 in `LsoBot-Config.psd1`:
    $dataFile = "$LsoScriptRoot\Logs\lsoBot-data.csv"

Add the following to line 736 in `LsoBot.ps1`:
    `Write-Output "$(Get-Timestamp), $Pilot, $Grade" | Out-file $dataFile -append`

Then restart the LSO BOT script. This will generate output to a file called `lsoBot-data.csv` in the `Logs/` folder of the application. 

TODO: Provide a patch file to apply to LSOBOT.

## Installing Dependencies

You need to install the following dependencies for GreenieBoard:

    pip install gspread
    pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib oauth2client 

TODO: Automate this using `pipenv`.

## Setting Up the Google Service Account

You will need a Google Cloud service account in order to programmatically access objects on Google Drive. Setting this up is outside the scope of this document. 

Currently, the file `token.js` is a placeholder for the real JSON credentials for accessing the service account. Use the Google Cloud dashboard to generate a file for your own service account.

## Creating the Sheets

GreenieBoard connects to two different Google Sheets to track data:
* **Feed sheet** - stores all of the landing events, including time, pilot, grade, comments, points, and wire. In addition to a master sheet listing all of the landings, the application will generate a separate worksheets per pilot. 
* **Greenie sheet** - contains the greenie board, with one worksheet per month. New worksheets will automatically be created when the month changes.

Both of these sheets will need a worksheet called "Template" which will be duplicated to generate new worksheets. For the feed sheet template, the format per column is `Time; Pilot; Grade; Comments; Wire; Points; Server`. The feed sheet will also need a worksheet called `Main Feed` where all of the landing events will be added.

For the greenie board template, this is a little more tricky since you may want to customize its look and feel. For the VF-111 greenie board, the empty slots (which is where landings will be logged) begins on column 8 (indexed from 1) and goes until column 43. The four slots before this is used for `Total Landings; Monthly Landings; Total Score Avg; Monthly Score Avg`. The index numbers are configured using the variables `emptySlotIndexStart` and `emptySlotIndexEnd` near the top of `GreenieBoard.py`.

Once you have these sheets created, you will need to (a) share them with the service account as 'Editor' (see above), and (b) update the sheet URLs in the file (see below).

## Configuring GreenieBoard

All of the configuration settings can be found near the top of `GreenieBoard.py`. 

## Running and Scheduling GreenieBoard

To test whether your setup is working, collect some traps into the CSV file and then run `GreenieBoard.py` (no command line arguments). If you did things correctly, you should see the two spreadsheets being populated with data.

Once everything is working, clear out landing events in Main Feed and delete the pilot worksheets. Also delete the monthly greenie boards in the greenie sheet. These will be repopulated once you run the scheduled app for the first time. Then schedule the program to be run regularly. Updating Google Sheets does take some time, so I would not recommend you run this more often than once per minute. In fact, to avoid coming up against request quota limits, you may want to run this at 5-minute intervals. 

If you do encounter quota limits from Google Cloud, please let me know. There may be further ways to optimize the code and/or inserting arbitrary delays to avoid this restriction.

## Pilot Naming Convention

Currently, GreenieBoard uses a hard-coded naming convention as follows: `Callsign | Modex | Squadron`. An example would be `Madgrim | 211 | VF-111`. If you want to customize this naming convention to your server, edit the function titled `parse_pilot()`.

TODO: Parametrize the name and squadron parsing.

## Acknowledgments and Bug Reports

GreenieBoard was created by Niklas 'Madgrim' Elmqvist, F-14A RIO in VF-111 (BelomarFleetfoot#0319 on Discord). Please direct bug reports to me.

The idea for this application, as well as the subject matter expertise, was provided by Flash, CO of VF-111, and Crash, CO of VA-196. Further feedback, bug reports, and suggestions have been provided by many other members of VF-111. 

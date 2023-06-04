# **KAM Contacts Clean-Up Automation Proposal**
### A scraper designed to get KAM Contacts information on LinkedIn.

---
## Installation
```python
pip install -r requirements.txt
```
Dependencies of the project are written in the requirements.txt

It was used with Python 3.9.13 

---
## Project Conditions:
1. have an IDE/text editor and confirm the installation of the project dependencies
2. parameters.py must have a working email with password
3. put the linkedin links on the sheet 1 of the linkedscrape.xlsm put them in column A only 
4. parameters.py and linkedinscrape.xlsm must be on the same directory as the project file 

## Usage
Run the program by clicking the run button or type on a command line:
```
python linkedinscrape.py
```
The script would open the linkedinscrape.xlsm, and the person would be prompted to enter First Row: and Last Row: 
(both referring to the excel file)  

Chrome would then open and direct you to linkedin, then logged in. 

NOTE: It is NOT NEEDED, but we recommend enabling Two-Factor Authenticator as an anti-bot detection 
hence, would need to TYPE IN THE COMMAND LINE the sent sms code.

The script would read the links, and go to each of them individually gettting certain information.

NOTE: Avoid using the automated browser, nor the linkedin account while the script is running because there would be a warning of bot behavior. 

The script finishes scraping after the chrome closed.

---
## Scraping Process System


After running the script it starts a session, opens the Google Chrome and go to LinkedIn. With the parameters python file that contains the user's info we can successfully, Log In. 

Next, the script would read through the Links starting from the prompted first row. It would then go to the profile, and scrape certain information. And once, the scraper got these it would store them on an excel workbook. 

It would then answer a conditional question, are there any more links? If yes, it would read that link, go to the profile, get information and store them into excel. It will answer no until there are no more links and reach the prompted last row. The scraper would then exit google chrome, quit the session and would print the results: Total no. of Profile Scraped and for how long was the session running.


---
## How the script see each profile

The script would get the following: Full Name, Location, Headline, Company name, and the current Job Title. In this example we have a screenshot of Patrice Torress's LinkedIn profile. The script would only look through two sections the Header, and Experience. 

The highlighted blue square indicates what the scraper is looking for, which is written in our script.

The header contains the Full Name, Location, and the Headline.

The experience section will contain the job title, and the company name. For the script to distinguish the current job title to other job titles, it would look on the word 'present’. If the word present is detected, it would get that job title.

---
## NOTE:

1. The progam can run on headless browser
2. The script cannot account for every possible profile permutations
3. It would be helpful to use a linkedin premium or head hunter account so that people wouldn't be notify when viewing their profile.
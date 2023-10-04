
## The project is intended to be an attempt to create an application that would be a tool for accountants to create financial reports from the data contained in the Trial Balance. Available reports are: Statement of Profit and Loss (SOPL) and Statement of Financial Position (SOFP).

![Picture1 of the Heroku mock terminal](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/Heroku mock terminal-Welcome.png)

- The program begins with a welcome message.
- Simple instructions printed on the terminal allow the user to enter the necessary data. 
- The program contains several algorithms for checking the correctness of data and the compliance of items in reports with items in TB as well.
- At the end, information about the report execution appears. 
- The data completed in the SOFP and SOPL reports can be viewed by the user in a Google sheet.
- Before clicking 'Run Program', the accountant prepares a blank report sheet.

![Picture of blank fs worksheet](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/blank SOFP sheet before calculations.png)



## Features

#### Retrieving GL codes from the user.

- The loop will repeatedly request data, until it is valid.
- Raises ValueError if strings cannot be found as the GL code or there are more than one position with the same GL code or if there aren't exactly 2 values provided by the user.


#### Checking the sum of a given column of numbers in the Trial Balance and collecting data for the financial statement.

- Selects data for individual reports.
- The financial statement spreadsheet is still blank.

![Picture of blank fs worksheet](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/blank SOPL sheet before calculations.png)


#### Prepares financial statement of a given type.

- Updates the worksheet with rounded values.
- Changes the sign of the numbers where and if required for given fs.
- Throws an error if any element in fs is not unique or there is no such element in TB.

![Picture of SOFP1](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/SOFP-1.png)

![Picture of SOFP2](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/SOFP-2.png)


#### Handle data in financial statement

- Adds a style to the sheet header.
- Calculates totals in columns for given groups of data.
- Computes Loss Before Tax and Loss for the Financial Period
- Updates the appropriate spreadsheet cells.


![Picture of SOPL](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/SOPL.png)


#### Final message

![Picture2 of the Heroku mock terminal](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/Heroku terminal-calc completed.png)



## Testing

- I have permanently conducting testing using Google API.
- I got "APIError: {'code': 429, 'message': "Quota exceeded...". This specific error is raised not always but often.
- Also I tested program asking an accountant for the review.



### Validator Testing


-  CI Python Linter validator did not detect any errors either

  ![screenshot](https://github.com/ireneuszcierpisz/love-accountancy/blob/main/media/CI Python Linter.png)



## Deployment

  - The app was deployed out to the web by Heroku. The steps to deploy are as follows:
  - From the Heroku dashboard click the “Create new app” button.
  - Name app.
  - Select region.
  - Click settings section.
  - Create config var as CREDS.
  - Add buildpacks as Python and Node.js
  - Select deploy section.
  - Choose deployment method (GitHub).
  - Search for your Github repository name.
  - Click "connect” to link up Heroku app to Github repository code.
  - Deploy using deploy branch option or automatic deploy.


## Credits

  - Project Love Accountancy is made thanks Code Institute lectures in Full Stack Software Development (E-commerce Applications) course.
  - The idea of ​​automating the preparation of financial statements from the data contained in the trial balance arose after conversations with accountants. In fact, this is what they do permanently. Thank you for your support and explaining the rules so that the application could be useful.

### Content

-




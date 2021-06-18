# TwitterScraper
### Uses [Twitter API](https://developer.twitter.com/en/docs/twitter-api) and [tweepy](https://www.tweepy.org/) to scrape data from profiles

* Authenticates the app by using the api credentials and a pin code
* Gets users using screen names 
* Scrapes the data related to users and saves them to an excel file

## Set UP
- install `pipenv`
    ```sh
    pip install pipenv
    ```
- install the required dependencies from the `pipfile`
    ```sh
    pipenv install
    ```

# Run
- Run the following command to execute the script
    ```sh
    pipenv run python twitter_scraper.py
    ```
- A browser window will pop up and will ask to authorize the app
- Please authorize and copy the authorization key to the terminal

## Tweaks
- Edit the variable `FILE_PATH` to change the path of input excel file
- Name the excel sheet with input urls as `Input`




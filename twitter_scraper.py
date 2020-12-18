import tweepy
import webbrowser
import time
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import json


# Dealing with Excel Files
FILE_NAME = 'Twitter Data.xlsx'
wb = load_workbook(FILE_NAME)

# Consumer API keys
with open('config.json') as f:
    CONFIG_DATA = json.load(f)
consumer_key = CONFIG_DATA.get('consumer_key')
consumer_secret_key = CONFIG_DATA.get('consumer_secret_key')
callback_uri = CONFIG_DATA.get('callback_uri')

# Authenticating the app
auth = tweepy.OAuthHandler(consumer_key, consumer_secret_key, callback_uri)
redirect_url = auth.get_authorization_url()

# opening the link with selenium
webbrowser.open(redirect_url)
user_pin_input = input('What is your pin? ')
auth.get_access_token(user_pin_input)

# calling the api  
api = tweepy.API(auth) 

# fetching the user 
def get_user_data(screen_name):
    user = api.get_user(screen_name) 
    
    # fetching the description 
    name = user.name
    followers_count = user.followers_count
    following_count = user.friends_count
    location = user.location
    verified = user.verified
    description = user.description 
    created_at = user.created_at.strftime('%d %B %Y')
    url = user.url


    data_dict = {
        'name': name,
        'followers_count': followers_count,
        'following_count': following_count,
        'description': description,
        'location': location,
        'url': url,
        'created_at': created_at,
        'verified': verified,
    }

    return data_dict


# Prepares the excel sheets and names the columns
def customize_excel_sheet():
    global wb
    output = wb.create_sheet('Output') if 'Output' not in wb.sheetnames else wb['Output']
    errors = wb.create_sheet('Errors') if 'Errors' not in wb.sheetnames else wb['Errors']
    
    font = Font(color="000000", bold=True)
    bg_color = PatternFill(fgColor='E8E8E8', fill_type='solid')

    # editing the output sheet
    output_column = zip(('A',  'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'), ('Screen Name', 'Name', 'Followers', 'Following', 'Description', 'Location', 'Link', 'Joined Date', 'Verified'))
    for col, value in output_column:
        cell = output[f'{col}1']
        cell.value = value
        cell.font = font
        cell.fill = bg_color
        output.freeze_panes = cell

        # fixing the column width
        output.column_dimensions[col].width = 20


# Generates the input links
def generate_screen_names():
    global wb
    inputs = wb['Input']
    for row in range(2, inputs.max_row + 1):
        # generates the links one by one
        if value := inputs[f"A{row}"].value:
            yield value


def insert_data_into_excel():
    customize_excel_sheet()
    for screen_name in generate_screen_names():
        user_data = get_user_data(screen_name)
        wb['Output'].append((
            screen_name, 
            user_data['name'], 
            user_data['followers_count'], 
            user_data['following_count'], 
            user_data['description'], 
            user_data['location'], 
            user_data['url'], 
            user_data['created_at'], 
            user_data['verified'],

        ))
    wb.save(FILE_NAME)


# Calling the main function
insert_data_into_excel()

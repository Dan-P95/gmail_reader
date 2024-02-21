This only works for a very specific use. A sample html file is included.

First, setup a virtual environment (.venv) and activate it
Then, in the .venv/bin/activate download the following using pip
(example command: 'python3 -m pip install google-api-python-client')
1. google-api-python-client
2. google-auth-oauthlib
3. beautifulsoup4
4. lxml
   
To try, open main and enter your own email at the top of the file next to "our_email ="

You must authorize the app to use the Gmail API. See this link on how to set it up https://developers.google.com/gmail/api/quickstart/python
Be sure to save the credentials in the root directory named 'credentials.json' 


Upon running the program the first time, it will open a new window to authorize the app. This will create a token.pickle file, the script can now be run.
See the output file in the Reservations folder, as well as the initial index.html file

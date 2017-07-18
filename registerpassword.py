import yagmail, keyring
#Leave yagmail, but replace sender email and password fields.
#Once this has been done, running this program once will permanantly 
#associate your email and password in the python keyring, unless of course
#it is run again with a different password, in which case it will overwrite 
#the previous.
keyring.set_password('yagmail', 'sender email', 'password')

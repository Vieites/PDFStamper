
from easygui import multpasswordbox
import keyring

import keyring
import win32timezone
import pickle
import os
import time

# User details: request the username and password and store them securely
def userdetails(systemname, msg, windowtitle, fieldnames):
    userfile = systemname + '.pkl'
    from easygui import multpasswordbox
    fieldvalues = []
    fieldvalues = multpasswordbox(msg, windowtitle, fieldnames)
    username = fieldvalues[0]
    userpassword =fieldvalues[1]
    f = open(userfile, 'wb')
    pickle.dump(username, f)
    f.close()
    keyring.set_password(systemname, username, userpassword)
    return fieldvalues



# Handle BTG network username and password
systemname = 'btgpdfstamper18'
userfile = systemname + '.pkl'
try:
    f = open(userfile, 'rb')
    btgusername = pickle.load(f)
    f.close
    btgpassword = keyring.get_password(systemname, btgusername)
except:
    windowtitle = 'BTG Network Credentials'
    msg = 'Enter your BTG credentials to access the BTG iQMS MasterControl (e.g. asouthall): '
    fieldnames = ["BTG Username", "BTG Password"]
    userdetails = userdetails(systemname, msg, windowtitle, fieldnames)
    btgusername = userdetails[0]
    btgpassword = userdetails[1]

# Handle BSC network username and password
systemname = 'bscpdfstamper18'
userfile = systemname + '.pkl'
try:
    f = open(userfile, 'rb')
    bscusername = pickle.load(f)
    f.close
    bscpassword = keyring.get_password(systemname, bscusername)
except:
    windowtitle = 'BSC Network Credentials'
    msg = 'Enter your BSC credentials to access BSC Teams (e.g.smithj): '
    fieldnames = ["BSC Username", "BSC Password"]
    userdetails = userdetails(systemname, msg, windowtitle, fieldnames)
    bscusername = userdetails[0]
    bscpassword = userdetails[1]

print(btgusername)
print(btgpassword)

print(bscusername)
print(bscpassword)

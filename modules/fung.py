import string
from os import listdir
from os.path import isfile, join

def format_filename(s):
    valid_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
    filename = ''.join(c for c in s if c in valid_chars)
    filename = filename.replace(' ','_') # I don't like spaces in filenames.
    return filename



onlyfiles = [f for f in listdir() if isfile(join("/home/farid/dev/python/synergy-gui/marked-data-sample/", f))]

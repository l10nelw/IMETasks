# listjob.py
# reads most recently modified cfg files and:
# - produces tab-separated-values for the harvest sheet's spiderjobs tab
# - produces list of domains
# - produces feedsettings.xml entries

import os
import glob
import configparser
import datetime

def sorted_dir_cfg(path):
    def mtime(file): return os.stat(os.path.join(path, file)).st_mtime
    return list(sorted(glob.glob(path + r'\*.cfg'), key=mtime))

DIR = os.environ['CFGBackup_Path'] # location of cfgfiles set in imetasks.ini
NUM = 5 # number of latest cfgfiles to pick up

FREQUENCY = { '30 minutes': '30mins', '1 hour': '1hr', '3 hours': '3hrs', '6 hours': '6hrs' }
FEEDSETTINGS = '<folder ID="{0}" NAME="{0}" OUTPUT="{0}.txt"><feed ID="{0}">{1}</feed></folder>\n'
DATE = datetime.date.today().isoformat()
USER = os.environ['USERNAME']

# values to take from config file
JOBITEMS = ( 'Directory', '//FeedURL', '//Frequency', '//Section', '//LastChange', 'URLFile', 'URL0', 'FixedFieldValue3', 'ImportMustHaveCSVs' )
# values to put into spiderjobs tab
ROWITEMS = ( 'type', 'spiderjob', 'url', 'section', 'source', 'urlfile', 'urlfnotes', 'fs' )

c = configparser.ConfigParser(allow_no_value=True)

# get list of NUM latest cfgfiles from DIR
cfglist = sorted_dir_cfg(DIR)[-NUM:]
cfglist.reverse()

# stuff to be output
errors = tabbedvalues = feedsettings = ''
domains = set()

for file in cfglist:
	with open(file, 'r') as f:
		# ignore anything outside sections
		content = ' '
		while content and content[0] != '[': content = f.readline() # keep reading lines until none left or '[' found
		if not content:
			errors += file + ' has no section headers.\n'
			continue # to next file
		content += f.read() # once a [section] is found, read rest of file
	try:
		c.read_string(content) # parse cfgfile
	except configparser.Error as e:
		errors += file + '\n'
		errors += str(e) + '\n'

if errors: errors += '*' * 100 + '\n'

job, row = dict(), dict()

for section in c.sections():
	if not c.options(section): continue # skip empty section
	
	for item in JOBITEMS:
		job[item] = c.get(section, item) if c.has_option(section, item) else ''
	
	if not job['Directory']: continue # skip non-spiderjob
	
	if job['URLFile']:
		row['type'] = 'rss'
		row['url'] = job['//FeedURL']
		row['fs'] = 'on'
		feedsettings += FEEDSETTINGS.format(section, job['//FeedURL'])
	else:
		row['type'] = FREQUENCY[job['//Frequency'].lower()] if job['//Frequency'] else ''
		row['url'] = job['URL0']
		row['fs'] = ''
	
	row['spiderjob'] = section
	row['section'] = job['//Section']
	row['urlfile'] = job['URLFile']
	row['source'] = job['FixedFieldValue3']
	row['urlfnotes'] = ''
	
	if job['//LastChange']:
		lastchange = job['//LastChange'].replace(' ', '\t', 2)
		lastchange = DATE + lastchange[lastchange.index('\t'):]
	else:
		lastchange = DATE + '\t' + USER
	
	for item in ROWITEMS:
		tabbedvalues += '\t' + row[item]
	tabbedvalues += '\ton\t' + lastchange + '\n'

	domains.add(job['ImportMustHaveCSVs'])

domains = '\n'.join(domains).replace('*','') + '\n'

with open('out.txt', 'w') as f:
	f.write(errors + tabbedvalues + domains + feedsettings)


from __future__ import division
import xml.etree.ElementTree as ET

import os
import sys

if (sys.version_info > (3, 0)):
	import tkinter as Tkinter 
	import tkinter.filedialog as tkFileDialog
	do_encode = True
else:
	import Tkinter
	import tkFileDialog
	do_encode = False

# Show the open file dialog
Tkinter.Tk().withdraw()
currdir = os.getcwd()
file_path = tkFileDialog.askopenfilename(initialdir=currdir, title='Please select the .msg file')

e = ET.parse(file_path).getroot()

messages = e.findall('Message')
total_count = len(messages)

output_template = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\' ?><smses count="{count}">{content}</smses>'
line_template = '<sms protocol="0" address="{address}" date="{timestamp:.0f}" type="{type}" subject="null" body="{body}" toa="null" sc_toa="null" service_center="null" read="1" status="-1" locked="0" />'
content = ''

print('Total count of messages: ' + str(total_count) + '\n')

i = 0
for m in messages:
	i += 1
	percentage = (i / total_count) * 100
	
	sys.stdout.write('\r')
	sys.stdout.write("[%-20s] processing %d" % ('=' * int(percentage / 5), i))
	sys.stdout.flush()
	
	# Message body
	text = m.find('Body').text
	if text is not None:
		body = text.replace("\"", "&quot;")
		if do_encode:
			body = body.encode('utf-8', 'ignore')
	else:
		# Fallback to empty string when Body is empty,
		# for some reason
		body = ''
	
	# Type --> 1=received, 2=sent
	type = '1' if m.find('IsIncoming').text == 'true' else '2'
	
	# Received message, get the sender
	if type == '1':
		address = m.find('Sender').text
	# Sent message, get one recipient
	elif type == '2':
		address = list(m.find('Recepients'))[0].text
	# Ouch, fallback
	else:
		address = ''
	
	if address is not None and address != '' and do_encode:
		address = address.encode('utf-8', 'ignore')
	
	# Uncomment and customize this for adding missing prefix
	# if address[0] != '+' and address[0].isdigit() and len(address) > 7:
	# 	address = '+39' + address
	
	# Parse the timestamp into UNIX milliseconds timestamp
	ts = int(m.find('LocalTimestamp').text) / 10000 - 11644473600000
	
	line = line_template.format(
		address=address,
		timestamp=ts,
		type=type,
		body=body
	)
	content += line

output = output_template.format(
	count=total_count,
	content=content
)

path = os.path.dirname(os.path.abspath(file_path))

with open(os.path.join(path, 'wp_messages.xml'), 'w') as file:
	file.write(output)

print('\n\nSuccess. Output to --> wp_messages.xml')

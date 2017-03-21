from __future__ import division
import xml.etree.ElementTree as ET

import os
import sys

if (sys.version_info >= (3, 0)):
	import tkinter as Tkinter
	import tkinter.filedialog as tkFileDialog
	isPy3 = True
else:
	import Tkinter
	import tkFileDialog
	isPy3 = False

# Show the open file dialog
Tkinter.Tk().withdraw()
currdir = os.getcwd()
file_path = tkFileDialog.askopenfilename(initialdir=currdir, title='Please select the .msg file')

e = ET.parse(file_path).getroot()

messages = e.findall('Message')
total_count = len(messages)

output_template = '<?xml version=\'1.0\' encoding=\'UTF-8\' standalone=\'yes\' ?><smses count="{count}">{content}</smses>'
sms_template = '<sms protocol="0" address="{address}" date="{timestamp:.0f}" type="{type}" subject="null" body="{body}" toa="null" sc_toa="null" service_center="null" read="{read}" status="-1" locked="0" />'
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
		if not isPy3:
			body = body.encode('utf-8', 'ignore')
	else:
		# Fallback to empty string when Body is empty,
		# for some reason
		body = ''

	# Type --> 1=received, 2=sent
	type = '1' if m.find('IsIncoming').text == 'true' else '2'

	# Read --> 0=no, 1=yes
	read = '1' if m.find('IsRead').text == 'true' else '0'

	# Received message, get the sender
	if type == '1':
		address = m.find('Sender').text
	# Sent message, get one recipient
	elif type == '2':
		address = list(m.find('Recepients'))[0].text
	# Ouch, fallback
	else:
		address = ''

	if address is not None and address != '' and not isPy3:
		address = address.encode('utf-8', 'ignore')

	# Uncomment and customize this for adding missing prefix
	# if address[0] != '+' and address[0].isdigit() and len(address) > 7:
	# 	address = '+39' + address

	# Parse the Windows file timestamp into UNIX milliseconds timestamp.
	ts = int(m.find('LocalTimestamp').text) / (10 * 1000 * 1000) - 11644473600

	line = sms_template.format(
		address=address,
		timestamp=ts,
		type=type,
		body=body,
		read=read
	)
	content += line

output = output_template.format(
	count=total_count,
	content=content
)

path = os.path.dirname(os.path.abspath(file_path))

if isPy3:
	with open(os.path.join(path, 'wp_messages.xml'), 'w', encoding='utf-8') as file:
		file.write(output)
else:
	with open(os.path.join(path, 'wp_messages.xml'), 'w') as file:
		file.write(output)

print('\n\nSuccess. Output to --> wp_messages.xml')

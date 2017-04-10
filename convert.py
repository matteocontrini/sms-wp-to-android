from __future__ import division
from base64 import b64decode
from cgi import escape
from datetime import datetime
from random import SystemRandom

import xml.etree.ElementTree as ET

import os
import string
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

# For hints on the attribute meanings see http://www.phonesdevelopers.info/1751299/
mms_template = '<mms text_only="0" ct_t="application/vnd.wap.multipart.related" using_mode="0" msg_box="{msg_box}" secret_mode="0" v="18" ct_cls="null" retr_txt_cs="null" d_rpt_st="0" favorite="0" deletable="0" sim_imsi="null" st="null" creator="com.github.matteocontrini.sms-wp-to-android" tr_id="{tr_id}" sim_slot="0" read="{read}" m_id="{m_id}" callback_set="0" m_type="{m_type}" locked="0" retr_txt="null" resp_txt="null" rr_st="0" safe_message="0" retr_st="null" reserved="0" msg_id="0" hidden="0" sub="null" seen="1" rr="129" ct_l="null" from_address="null" m_size="null" exp="{exp}" sub_cs="null" sub_id="-1" app_id="0" resp_st="{resp_st}" date="{date:.0f}" date_sent="0" pri="129" address="{address}" d_rpt="129" d_tm="null" read_status="null" device_name="null" spam_report="0" rpt_a="null" m_cls="personal" readable_date="{readable_date}"><parts>{parts}</parts><addrs>{addrs}</addrs></mms>'
mms_part_template = '<part seq="{seq}" ct="{ct}" name="null" chset="{chset}" cd="null" fn="null" cid="&lt;{cid}&gt;" cl="{cl}" ctt_s="null" ctt_t="null" text="{text}" {data} />'
mms_addr_template = '<addr address="{address}" type="{type}" charset="{charset}" />'

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
		body = text.replace("\n", "&#10;")
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

	sender = m.find('Sender').text
	recepients = m.find('Recepients')
	if len(recepients) > 0:
		recepient = recepients[0].text
	else:
		recepient = 'null'

	# Received message, get the sender
	if type == '1':
		address = sender
	# Sent message, get one recipient
	elif type == '2':
		address = recepient
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

	attachments = m.find('Attachments')
	if attachments is None:
		line = sms_template.format(
			address=address,
			timestamp=ts,
			type=type,
			body=body,
			read=read
		)
	else:
		UTF8 = '106'

		parts = ''
		for attachment in attachments.findall('MessageAttachment'):
			content_type = attachment.find('AttachmentContentType').text

			if content_type == 'application/smil':
				seq = '-1'
				character_set = 'null'
				content_id = 'smil'
				content_location = 'smil{0:05d}.xml'.format(i)
				text = 'null'
				data = ''
			elif content_type == 'text/plain':
				seq = '0'
				character_set = UTF8
				content_id = 'text'
				content_location = 'text{0:05d}.txt'.format(i)
				unicode_bytes = b64decode(attachment.find('AttachmentDataBase64String').text)
				text = escape(unicode_bytes.decode('utf-16')).encode('utf-8')
				data = ''
			elif content_type == 'image/jpeg':
				seq = '0'
				character_set = 'null'
				content_id = 'image'
				content_location = 'image{0:05d}.jpg'.format(i)
				text = 'null'
				data = attachment.find('AttachmentDataBase64String').text
			elif content_type == 'image/png':
				seq = '0'
				character_set = 'null'
				content_id = 'image'
				content_location = 'image{0:05d}.png'.format(i)
				text = 'null'
				data = attachment.find('AttachmentDataBase64String').text
			elif content_type == 'text/x-vCard':
				seq = '0'
				character_set = 'null'
				content_id = 'vCard'
				content_location = 'vCard{0:05d}.xml'.format(i)
				text = 'null'
				data = attachment.find('AttachmentDataBase64String').text
			else:
				seq = '0'
				character_set = 'null'
				content_id = 'unknown'
				content_location = 'unknown{0:05d}'.format(i)
				text = 'null'
				data = attachment.find('AttachmentDataBase64String').text

			# Really only write the "data" attribute if there is data defined.
			if len(data) > 0:
				data = 'data="{0}"'.format(data)

			parts += mms_part_template.format(
				seq=seq,
				ct=content_type,
				chset=character_set,
				cid=content_id,
				cl=content_location,
				text=text,
				data=data
			)

		TYPE_FROM = '137'
		TYPE_TO = '151'
		addrs = mms_addr_template.format(address = sender, type = TYPE_FROM, charset = UTF8)
		addrs += mms_addr_template.format(address = recepient, type = TYPE_TO, charset = UTF8)

		id_set = string.ascii_letters + string.digits

		N = 12 # Somewhat arbitrary.
		transaction_id = ''.join(SystemRandom().choice(id_set) for _ in range(N))

		N = 26 # Somewhat arbitrary.
		message_id = ''.join(SystemRandom().choice(id_set) for _ in range(N))

		MESSAGE_TYPE_SEND_REQ = 128
		MESSAGE_TYPE_RETRIEVE_CONF = 132
		message_type = MESSAGE_TYPE_RETRIEVE_CONF if type == '1' else MESSAGE_TYPE_SEND_REQ

		expiry = 'null' if type == '1' else '604800'
		response_status = 'null' if type == '1' else '128'

		# Get a readable date with abbreviated month and non-zero padded day of month.
		readable_date = '{dt:%b} {dt.day}, {dt.year} {dt:%H}:{dt:%M}:{dt:%S}'.format(dt = datetime.fromtimestamp(ts))

		line = mms_template.format(
			msg_box=type,
			tr_id=transaction_id,
			read=read,
			m_id=message_id,
			m_type=message_type,
			exp=expiry,
			resp_st=response_status,
			date=ts,
			address=address,
			readable_date=readable_date,
			parts=parts,
			addrs=addrs
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

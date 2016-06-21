#!/usr/bin/env python
#
# outlook_reaper.py
# Extract emails as mbox.
#
# (c) Copyright James Aylett 2008
# License: BSD
# Technique based on <http://www.boddie.org.uk/python/COM.html>,
# resolve prereqs from there before running.
#
# You need to provide a map from LDAP-style (typically Exchange)
# addresses to email addresses, or some of your headers are going to
# be wrong. Run it over the folder and it'll report unmapped addresses.
#
# FIXME: doesn't cope with .msg attachments (should unpack to
# message/rfc-822 but just includes the Outlook .msg binary as
# application/octet-stream). RDO (Redemption) is supposed to help with
# this, but I can't get it to initialise because the sessions stuff
# isn't working for some reason.

LDAP_MAP = {
}

import sys, email, win32com.client, os, traceback
import traceback, mimetypes
from email.mime.base import MIMEBase
from email.mime.message import MIMEMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.audio import MIMEAudio
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import email.charset

rdo = None

class MyException(Exception):
    def __init__(self, m):
	self.m = m

    def __str__(self):
	return self.m

ldap_stack = []
def fixup(addr):
    if len(addr)>1 and addr[0]==u'/':
	# LDAP-style
	if LDAP_MAP.has_key(addr.strip()):
	    return LDAP_MAP[addr.strip()]
	if addr not in ldap_stack:
	    ldap_stack.append(addr.strip())
    return addr.strip()

# This probably still isn't quite right, but doesn't actually crash
# (and seems to work okay, although difficult to be sure as the bodies
# I was having trouble with now end up bas64 encoded; I need to test
# properly at some point).
local_convert = email.charset.Charset.convert
def better_convert(self, s):
    try:
	if type(s) is unicode:
	    t = s.encode('ascii')
	    t = None
	return local_convert(self, s)
    except UnicodeEncodeError:
	return local_convert(self, s.encode('utf-8'))
email.charset.Charset.convert = better_convert

def write_message(f, m):
    ms = build_message(m)
    if ms!=None:
        f.write(str(ms))
        f.write('\n\n')

def build_message(m):
    try:
	if m.InternetCodepage==65001:
	    charset = 'utf-8'
	elif m.InternetCodepage==28591:
	    charset = 'iso-8859-1'
	elif m.InternetCodepage==1252:
	    charset = 'iso-8859-1' # safer and frankly more sane
	elif m.InternetCodepage==20127:
	    charset = 'us-ascii'
	else:
	    charset = 'iso-8859-1'
    except AttributeError:
	# missing InternetCodepage, so (probably) not a message
	return None

#    if charset=='utf-8':
#	charset='iso-8859-1'
#	# for some reason the weird email charset stuff doesn't get this
#	# right and tries to output utf-8 using (I think) base64 and the
#	# ascii codec, which doesn't work right; we promote if we need to
#
#    for cs in ['utf-8', 'iso-8859-1']:
#	c = email.charset.Charset(cs)
#	for at in ['input_charset', 'header_encoding', 'body_encoding', 'output_charset', 'input_codec', 'output_codec']:
#	    print "Charset(%s).%s = %s" % (cs, at, getattr(c,at))
#    sys.exit(0)

    try:
	t = m.Body.encode(charset)
	t = m.HTMLBody.encode(charset)
	t = None
    except UnicodeEncodeError:
	if charset!='iso-8859-1':
	    #print "Given charset (%s) lies, promoting to iso-8859-1" % charset
	    charset = 'iso-8859-1'
	else:
	    #print "Given charset (%s) lies, promoting to utf-8" % charset
	    charset = 'utf-8'

    if m.BodyFormat==win32com.client.constants.olFormatPlain:
	b = MIMEText(m.Body, 'plain', charset)
    elif m.BodyFormat==win32com.client.constants.olFormatHTML:
	try:
	    bt = MIMEText(m.Body, 'plain', charset)
	    bh = MIMEText(m.HTMLBody, 'html', charset)
	    b = MIMEMultipart('alternative')
	    b.attach(bt)
	    b.attach(bh)
	except:
	    print "Problem with charset=%s, message=%s" % (charset, m.Subject)
	    raise
    elif m.BodyFormat==win32com.client.constants.olFormatRichText:
	b = MIMEText(m.Body, 'enriched', charset)
    else:
	raise MyException("Unknown body format %i" % m.BodyFormat)

    if len(m.Attachments)>=1:
	b2 = MIMEMultipart('mixed')
	b2.attach(b)
	b = b2
    #mess = MIMEMessage(b)
    mess = b
    
    mess['From'] = "%s <%s>" % (m.SenderName, fixup(m.SenderEmailAddress))
    rs = []
    for ri in range(1, len(m.Recipients)+1):
	r = m.Recipients[ri]
	rs.append( "%s <%s>" % (r.Name.strip(), fixup(r.Address)) )
    if len(rs)>0:
	mess['To'] = "%s" % ','.join(rs)
    mess['Subject'] = "%s" % m.Subject.strip()
    mess['Date'] = ("%s" % m.SentOn.Format('%a, %d %b %Y %H:%M:%S +0000')).strip()
    try:
	rs = []
	for ri in range(1, len(m.ReplyRecipients)+1):
	    r = m.ReplyRecipients[ri]
	    rs.append( "%s <%s>" % (r.Name.strip(), fixup(r.Address)) )
	if len(rs)>0:
	    mess['Reply-To'] = "%s" % ','.join(rs)
    except:
	print "Problems with ReplyRecipients"
        traceback.print_exc()
    # These probably don't have email addresses, wargh
    if m.CC!='':
	mess['Cc'] = m.CC.strip()
    if m.BCC!='':
	mess['BCC'] = m.BCC.strip()
    # m.UserProperties (1-indexed list, .Name, .Value)
    for pi in range(1, len(m.UserProperties)+1):
	p = m.UserProperties[pi]
	mess['User-%s' % p.Name.strip()] = str(p.Value).strip()
    #for pi in range(1, len(m.ItemProperties)+1):
#	p = m.ItemProperties[pi]
#	mess['Item-%s' % p.Name.strip()] = str(p.Value).strip()

    for ai in range(1, len(m.Attachments)+1):
        try:
            at = m.Attachments[ai]
            if at.Type==win32com.client.constants.olEmbeddeditem:
                # embedded message, recursively parse it
                # we can only do this if RDO is installed (well, we could
                # save out and create item from template, but that will
                # mess around with some of the headers, so it's better to
                # preserve the original binary .msg if RDO isn't available)
                # print "found .msg attachment"
                msg = None
                try:
                    msg = rdo.GetMessageFromID(m.EntryID)
                    print "extracting via rdo (id=%s)" % str(m.EntryID)
                    at2 = msg.Attachments[ai]
                    msg = at2.EmbeddedMsg
                    msg = None
                    at2 = None
                    b.attach(build_message(msg))
                    continue
                except:
                    # rdo is None, or something else went wrong
                    if msg:
                        print "failed to extract via rdo, including .msg as binary"
                    pass

            fn = os.tmpnam()
            at.SaveAsFile(fn)
            fa = file(fn, 'rb')
            fac = fa.read()
            fa.close()
            os.unlink(fn)
            typ = mimetypes.guess_type(at.FileName) # urgh
            if typ==None:
                typ='application/octet-stream'
            if type(typ) is tuple or type(typ) is list:
                typ = typ[0]
            try:
                typ = typ.split('/')
            except:
                typ = ['application', 'octet-stream']
            if typ[0]=='application':
                ma = MIMEApplication(fac, typ[1])
            elif typ[0]=='image':
                ma = MIMEImage(fac, typ[1])
            elif typ[0]=='audio':
                ma = MIMEAudio(fac, typ[1])
            elif typ[0]=='text':
                ma = MIMEText(fac, typ[1], 'utf-8')
            elif typ[0]=='video':
                ma = MIMEApplication(fac, typ[1])
                ma.set_type('/'.join(typ))
            elif typ[0]=='message' and typ[1]=='rfc822':
                ma = MIMEApplication(fac, typ[1])
                ma.set_type('/'.join(typ))
            else:
                raise MyException("Gah, failed on %s[%s]" % (at.FileName, '/'.join(typ)))
            ma.add_header('Content-Disposition', 'attachment', filename=at.FileName)
            b.attach(ma)
        except:
            # couldn't add that attachment, just skip it...
            try:
                print "Skipping attachment %i for %s" % (ai, m.Subject.strip())
            except:
                pass
            traceback.print_exc()
            pass
    return mess

def main(outf):
    global rdo
    outlook = win32com.client.Dispatch("Outlook.Application")
    context = outlook.Session
    context.Logon()

    try:
	rdo = win32com.client.Dispatch("Redemption.RDOSession")
        rdo.Logon()
        print "RDO is installed: this should all work better."
    except:
        traceback.print_exc()
	rdo = None

    while 1:
	print "-" * 60
	try:
	    print context.Name
	except:
	    print "[Root]"
	subs = []
	try:
	    folders = context.Folders
	    for i in range(1, len(folders)+1):
		try:
		    subs.append(folders[i])
		except:
		    print "folders[%i] funted" % i
	except:
	    print "No Folders property?"
	for i in range(0, len(subs)):
	    print i, subs[i].Name
	print "Select number or [E]xtract [Q]uit, followed by <Return>"
	s = raw_input().strip().upper()[0]
	if s=='E':
            print "Extracting to %s" % outf
	    f = file(outf, 'a')
	    try:
		items = context.Items
		for i in range(1, len(items)+1):
		    try:
			write_message(f, items[i])
		    except KeyboardInterrupt:
			raise
		    except SystemExit:
			raise
		    except:
			traceback.print_exc()
			#print sys.exc_info()[1]
	    except KeyboardInterrupt:
		raise
	    except SystemExit:
		raise
	    except:
		print "No Items property?"
                traceback.print_exc()
	    f.close()
	    sys.exit(0)
	elif s=='Q':
	    sys.exit(0)
	else:
	    try:
		i = int(s)
		context = subs[i]
	    except:
		print "Oops"

if __name__=="__main__":
    try:
	main(sys.argv[1])
    except SystemExit:
	if len(ldap_stack)>0:
	    print "Exiting: unresolved LDAP addresses:"
	    for ldap_addr in ldap_stack:
		print ldap_addr
	    print
	sys.exit(0)

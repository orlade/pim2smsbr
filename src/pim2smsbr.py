'''
A simple script to convert a PIM Backup file into an XML file usable by 
SMS Backup & Restore.
Created on 30/06/2011

@author: Oliver Lade
@contact: http://piemaster.net/
@contact: piemaster21@gmail.com

@license: Simplified BSD License

Copyright (c) 2011, Oliver Lade
All rights reserved.

Redistribution and use in source and binary forms, with or without modification,
are permitted provided that the following conditions are met:

 - Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.
 - Redistributions in binary form must reproduce the above copyright notice, 
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'''

import csv
import os
import re
import sys
import time
from zipfile import ZipFile
from xml.sax.saxutils import escape

SMS_LABEL = "IPM.SMStext"
DETAIL_URL = "http://piemaster.net/tools/winmo-android-sms-converter/"

# Create a CSV reader for the source file
def get_reader(source):
    csv_file = None
    ext = os.path.splitext(source)[1].lower()
    
    # If the message file is given directly, open it
    if ext in ('.csm', '.csv'):
        csv_file = open(source, 'r')
    # If given directly but from binary backup, error
    elif ext == '.pbm':
        throw_binary_error()
    # If the PIB file is given, extract the message file and open it
    elif ext == '.pib':
        zip = ZipFile(source, 'r')
        for filename in zip.namelist():
            ext = os.path.splitext(filename)[1].lower()
            if ext == '.csm':
                csv_file = zip.open(filename, 'r')
                break
            # Evidence of binary backup
            elif ext == '.pbm':
                throw_binary_error()
    else:
        display_error("Unknown input file type '%s', please use the original .pib or .csm backup file." % ext)

    if not csv_file:
        display_error("Couldn't find messages file. Please check your input path and contents.")
        
    # Read the file contents
    lines = csv_file.read().decode('utf-16', 'replace').encode('utf-8').splitlines()
    sms_reader = csv.reader(lines, delimiter=';', quotechar='"', escapechar='\\')
    # Return the reader and the eventual source filename
    return sms_reader

# Convert the PIM file to XML
def convert(source, out):
    start_time = time.time()
    print " - Reading input from '%s'..." % source
    try:
        sms_reader = get_reader(source)

    except IOError:
        display_error("Input file not found at '%s', aborting" % source)

    print " - Processing SMS messages"
    print "-"*40
    print " - Working..."

    items = []
    line_num = 0
    sms_count = 0
    warn_count = 0
    # For each message, using the iterator manually
    try:
        while 1:
            line_num += 1
            # Try to read the row first for encoding errors
            try:
                row = sms_reader.next()
            except UnicodeEncodeError:
                display_warning("Failed to decode line (UnicodeEncodeError), skipping...", line_num)
                warn_count += 1
                continue
            except csv.Error:
                display_warning("Failed to decode line (CSV error), skipping...", line_num)
                warn_count += 1
                continue

            # Process the contents of the row
            if not row:
                continue

            try:
                msg_class = row[10]
            except IndexError:
                display_warning("Line incorrectly formed, skipping...", line_num)
                warn_count += 1
                continue

            # If the entry is an SMS
            if msg_class.lower() == SMS_LABEL.lower():
                # Process it
                items.append(process(row, line_num))
                sms_count += 1

    except StopIteration:
        print "-"*40
        print " - Processing of %d messages complete!" % sms_count
        if warn_count > 0:
            print ""
            message = '''%d warnings generated.
 You may wish to correct these warnings manually in the source file (remember 
 to create a backup copy first), or for further assistance, see below.
 
 More information   - %s
 Help in the forums - http://piemaster.net/forums/''' % (warn_count, DETAIL_URL)
            display_warning(message)

    # If an output file path was not specified, generate from source path
    if not out:
        out = source[:-4] + '.xml'

    # Generate and write the XML output
    print " - Writing output to '%s'..." % out

    # Sort the list of output items by date
    from operator import itemgetter
    item_list = sorted(items, key=itemgetter('date'))
    # Build the output string
    out_str = ''
    for item in item_list:
        out_str += '\t%s\n' % item_to_xml(item)
    # Write the output string
    out_file = open(out, 'w')
    out_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n')
    out_file.write('<smses count="%d">\n%s</smses>' % (sms_count, out_str))

    # Finish up
    time_taken = time.time() - start_time
    print " - Output written"
    print " - Conversion complete! (%.2f secs)" % time_taken

# Process one row of the PIM file to a dictionary of XML field strings
def process(row, line_num=None):
    item = {}
    item['subject'] = row[4]
    item['body'] = escape(row[5]).replace('"', '&quot;').rstrip("- GSM")
    # If sender address is empty, then message was sent, not received
    was_sent = (row[2] == '')

    item['date'] = 0
    if row[16] != '':
        # Date should be timestamp in ms
        item['date'] = int(time.mktime(time.strptime(row[16], '%Y,%m,%d,%H,%M,%S'))) * 1000

    # If the message was sent
    if was_sent:
        item['type'] = 2
        try:
            item['address'] = row[18].split(';')[2].strip('\\')
        except:
            display_warning("Sent message destination not found, leaving empty...", line_num)
            item['address'] = ''

    # If the SMS was received
    else:
        item['type'] = 1
        # Match a string of digits with an optional plus
        match = re.search('[^+\d]*(\+*[\d]+)\D*', row[2])
        item['address'] = match.group(1) if match else None

    return item

# Generate the XML node from a dictionary of items generated in process()
def item_to_xml(item):
    return '<sms protocol="%s" address="%s" date="%s" type="%s" subject="%s" body="%s" toa="%s" sc_toa="%s" service_center="%s" read="%s" status="%s" locked="%s" />' % \
    (0, item['address'], item['date'], item['type'], item['subject'], item['body'], 'null', 'null', 'null', 1, 0, 0)
    
# Throw a warning to the user and continue
def display_warning(message, line=None):
    if line:
        print "WARNING (line %s): %s" % (line, message)
    else:
        print "WARNING: %s" % message
    
# Throw an error to the user and exit peacefully
def display_error(message):
    print '''
 !!!!!!!!!
 ! ERROR !
 !!!!!!!!!
 
 -> %s
 
 For help resolving this error:
 More information   - %s
 Help in the forums - http://piemaster.net/forums/''' % (message, DETAIL_URL)
    sys.exit()
    
# Throw an error informing the user to export uncompressed
def throw_binary_error():
    message = """Couldn't load binary messages file.
 Please ensure you DISABLE the BINARY BACKUP option in PIM Backup."""
    display_error(message)
    
    
if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description="Convert a PIM Backup messages file to SMS Backup & Restore-compatible XML.")
    parser.add_argument('source', metavar='source_file', type=str, nargs=1,
                       help="the source file to convert")
    parser.add_argument('out', metavar='out_file', type=str, nargs='?',
                       help="optionally specify a file to write the output to")
    args = parser.parse_args()

    # Print some contact details
    print 
    print "-"*40
    print "PIM Backup to SMS Backup & Restore Converter"
    print " Written by Oliver Lade (piemaster21@gmail.com)"
    print " More information at %s" % DETAIL_URL
    print " Questions and comments very welcome!"
    print "-"*40
    print 
    
    # Convert the given input file
    convert(args.source[0], args.out)

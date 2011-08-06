'''
A simple script to convert a PIM Backup file into an XML file usable by SMS Backup & Restore.
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

# Create a CSV reader for the source file
def get_reader(source):
    csv_file = None
    ext = os.path.splitext(source)[1].lower()
    # If the message file is given directly, open it
    if ext.lower() in ('.csm', '.csv'):
        csv_file = open(source, 'r')
    # If the PIB file is given, extract the message file and open it
    elif ext.lower() == '.pib':
        zip = ZipFile(source, 'r')
        for filename in zip.namelist():
            if os.path.splitext(filename)[1].lower() == '.csm':
                csv_file = zip.open(filename, 'r')
                break
    else:
        print "ERROR: Unknown input file type '%s', please use the original .pib or .csm backup file." % ext
        sys.exit()

    if not csv_file:
        print "ERROR: Couldn't load messages file, please check your input path and contents."
        sys.exit()

    # Read the file contents
    sms_text = csv_file.read().decode('utf-16').split(os.linesep)
    sms_reader = csv.reader(sms_text, delimiter=';', quotechar='"', escapechar='\\')
    # Return the reader and the eventual source filename
    return sms_reader

# Convert the PIM file to XML
def convert(source, out):
    start_time = time.time()
    out_str = ''
    print " - Reading input from '%s'..." % source
    try:
        sms_reader = get_reader(source)

    except IOError:
        print "ERROR: Input file not found at '%s', aborting" % source
        sys.exit()

    print " - Processing SMS messages"
    print "----------------------------------------"
    print " - Working..."

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
                print "WARNING (line %d): Failed to decode line, skipping..." % line_num
                warn_count += 1
                continue

            # Process the contents of the row
            if not row:
                continue

            try:
                msg_class = row[10]
            except IndexError:
                print "WARNING (line %d): Line incorrectly formed, skipping..." % line_num
                warn_count += 1
                continue

            # If the entry is an SMS
            if msg_class.lower() == SMS_LABEL.lower():
                # Process it
                msg_xml = process(row)
                out_str += '\t%s\n' % msg_xml
                sms_count += 1

    except StopIteration:
        print "----------------------------------------"
        print " - Processing of %d messages complete!" % sms_count
        if warn_count > 0:
            print "\nWARNING: %d warnings generated." % warn_count
            print "You may like to review the contents of the given lines"
            print "in the source file and correct them manually."
            print "(If source is .pib, change to .zip and extract .csm file)\n"

    # If an output file path was not specified, generate from source path
    if not out:
        out = source[:-4] + '.xml'

    # Write the output
    print " - Writing output to '%s'..." % out
    out_file = open(out, 'w')
    out_file.write('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n')
    out_file.write('<smses count="%d">\n%s</smses>' % (sms_count, out_str))

    time_taken = time.time() - start_time
    print " - Output written"
    print " - Conversion complete! (%.2f secs)" % time_taken

# Process one row of the PIM file to an XML node
def process(row):
    subject = row[4]
    body = escape(row[5]).replace('"', '&quot;')
    # If sender address is empty, then message was sent
    was_sent = (row[2] == '')
    
    date = 0
    if row[16] != '':
        # Date should be timestamp in ms
        print row[16]
        date = int(time.mktime(time.strptime(row[16], '%Y,%m,%d,%H,%M,%S'))) * 1000

    # If the SMS was received (sender not empty)
    if was_sent:
        # Match a string of digits with an optional plus
        match = re.search('[^+\d]*(\+*[\d]+)\D*', row[2])
        address = match.group(1) if match else None
        type = 1
        status = -1

    # Else the message was sent
    else:
        try:
            address = row[18].split(';')[2].strip('\\')
        except:
            address = ''
        type = 2
        status = 0

    # Generate the XML node
    msg_xml = '<sms protocol="%s" address="%s" date="%s" type="%s" subject="%s" body="%s" toa="%s" sc_toa="%s" service_center="%s" read="%s" status="%s" locked="%s" />' % \
    (0, address, date, type, subject, body, 'null', 'null', 'null', 1, status, 0)

    return msg_xml

if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser(description="Convert a PIM Backup messages file to SMS Backup & Restore-compatible XML.")
    parser.add_argument('source', metavar='source_file', type=str, nargs=1,
                       help="the source file to convert")
    parser.add_argument('out', metavar='out_file', type=str, nargs='?',
                       help="optionally specify a file to write the output to")
    args = parser.parse_args()

    # Print some contact details
    print "PIM Backup to SMS Backup & Restore Converter"
    print " Written by Oliver Lade (piemaster21@gmail.com)"
    print " More information at http://piemaster.net/tools/winmo-android-sms-converter/"
    print " Questions and comments very welcome!\n"

    # Convert the given input file
    convert(args.source[0], args.out)

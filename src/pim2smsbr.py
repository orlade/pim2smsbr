'''
Created on 30/06/2011

@author: Oliver Lade
@contact: http://piemaster.net/
@contact: piemaster21@gmail.com
'''
import csv
import time
import re
import os
from zipfile import ZipFile
from xml.sax.saxutils import escape

# NOTE: The PIM Backup file must first be manually converted to "UTF-8 without BOM" encoding

# Create a CSV reader for the source file
def get_reader(source):
    ext = source[-4:]
    # If the message file is given directly, open it
    if ext == ".csm" or ext == ".csv":
        csv_file = open(source, "rb")
    # If the PIB file is given, extract the message file and open it
    elif ext == ".pib":
        zip = ZipFile(source, 'r')
        for filename in zip.namelist():
            if filename[-4:] == ".csm":
                csv_file = zip.open(filename, 'r')
                break
    else:
        raise Exception("Unknown input file type, please use .pib or .csm.")
    
    if not csv_file:
        raise Exception("Couldn't load messages file, please check your input.")
    
    # Read the file contents
    sms_text = csv_file.read().decode("utf-16").split(os.linesep)
    sms_reader = csv.reader(sms_text, delimiter=';', quotechar='"', escapechar="\\")
    # Return the reader and the eventual source filename
    return sms_reader
    
# Convert the PIM file to XML
def convert(source, out):
    out_str = ""
    print "Reading input from " + source + "..."
    sms_reader = get_reader(source)
    
    print "Processing SMS messages..."
    sms_count = 0
    # For each message
    for row in sms_reader:
        if not row:
            continue
        msg_class = row[10]
        # If the entry is an SMS
        if msg_class == "IPM.SMStext":
            # Process it
            msg_xml = process(row)
            out_str += "\t" + msg_xml + "\n"
            sms_count += 1
    print "Processing %d messages complete!" % sms_count

    # If an output file path was not specified, generate from source path
    if not out:
        out = source[:-4] + '.xml'

    # Write the output
    print "Writing output to " + out
    out_file = open(out, "wb")
    out_file.write("<?xml version='1.0' encoding='UTF-8' standalone='yes' ?>\n")
    out_file.write('<smses count="%d">\n%s</smses>' % (sms_count, out_str))
    print "Output written!"

# Process one row of the PIM file to an XML node
def process(row):
    subject = row[4]
    body = escape(row[5]).replace('"', '&quot;')
    date = 0;
    if row[16] != "":
        date = int(time.mktime(time.strptime(row[16], "%Y,%m,%d,%H,%M,%S")))

    # If the SMS was received (sender not empty)
    if row[2]:
        # Match a string of digits with an optional plus
        match = re.search("[^+\d]*(\+*[\d]+)\D*", row[2])
        address = match.group(1) if match else None
        type = 1
        svc_ctr = "+491722270333"
        status = -1
        
    # Else the message was sent
    else:
        address = row[18].split(';')[2].strip('\\')
        type = 2
        svc_ctr = "null"
        status = 0

    # Generate the XML node
    msg_xml = '<sms protocol="%s" address="%s" date="%s" type="%s" subject="%s" body="%s" toa="%s" sc_toa="%s" service_center="%s" read="%s" status="%s" locked="%s" />' % \
    (0, address, date, type, subject, body, "null", "null", svc_ctr, 1, status, 0)

    return msg_xml

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description='Convert a PIM Backup messages file to SMS Backup & Restore-compatible XML.')
    parser.add_argument('source', metavar='source_file', type=str, nargs=1,
                       help='the source file to convert')
    parser.add_argument('out', metavar='out_file', type=str, nargs='?',
                       help='optionally specify a file to write the output to')
    args = parser.parse_args()

    # Convert the given input file
    convert(args.source[0], args.out)

#
# Name:			jsonreports.py
# Author:		Julio Carranza
# Date/Version:	11-30-2017 first version
#
#
#
# Description:	Opens an excel file and reads cells in a column
#               specifying urls for JSONs files. It will create
#               a directory and download JSON files to it. It
#               creates the JSON file based on the URL name.
#               Then it reads the duration of the duration field.
#               If it can't read it, it will declare the JSON is
#               malformed. Finally, it creates a new excel file,
#               using the initial name + timestamps.

# usage:        python jsonreports.py jsontracker.xlsx
#


from openpyxl import load_workbook
import sys, os, re, json, time
from urllib2 import urlopen

if len(sys.argv) != 2:
    print "ERROR: enter one argument specifying excel file with json links"
    exit()
else:
    trackerfile = sys.argv[1:]
    trkrfl = trackerfile[0]

timestr = time.strftime("%Y%m%d%H%M")


wb = load_workbook( filename = trkrfl )
trkrfl = trkrfl.split('.xlsx')
jsondirectory = trkrfl[0]

if not os.path.exists(jsondirectory):
    print "Creating directory", jsondirectory
    os.makedirs(jsondirectory)

resultssheet = wb['Sheet1']



#intial row
a=10
#last row to check
z=250
#column with urls
urlcolumn = 'V'
#column to write JSON filename
scolumn = 'S'
#column to write durationtime
durcolumn = 'U'



while (a < z):
    celdaV = urlcolumn + str(a)
    jsonurlcell = resultssheet[celdaV].value
    celdaS = scolumn + str(a)
    celdaU = durcolumn + str(a)

    jsontosavecell = ''
    durationcell = ''

    if (jsonurlcell is not None):

        regexx4RXurl = '([rR]+[1]?[0-9]:\s*http://[^\s]*/JT_[^\s]*.pl_[0-9]+/ezLog/Tc[0-9]+[^\s]*/report.json)'
        regexx4url = '^http://.*/JT_(.*).pl(_[0-9]+)/ezLog/(Tc[0-9]+[^\s]*)/report.json'
        regexx4fl = '([rR]+[1]?[0-9]):\s*http://.*/JT_([^\s]*).pl(_[0-9]+)/ezLog/(Tc[0-9]+[^\s]*)/report.json'
        if re.match( regexx4RXurl, jsonurlcell):

            jsonslist = re.findall(regexx4RXurl, jsonurlcell)
            print jsonurlcell
            for rtr in jsonslist:


                RX = re.match( regexx4fl, rtr).group(1)
                jsontosave = RX +'_'+ re.match( regexx4fl, rtr).group(2) + re.match( regexx4fl, rtr).group(3) +'_'+ re.match( regexx4fl, rtr).group(4) +'.json'
                jsontosavecell += RX + ': '+ jsontosave + ' '

                dirjsontofile = jsondirectory + '/'+ jsontosave
                jsonurl = re.match( '.*(http://.*report.json).*', rtr).group(1)
                try:
                    fp = open(dirjsontofile, 'wb')
                    req = urlopen(jsonurl)
                    for line in req:
                        fp.write(line)
                    fp.close()
                    with open(dirjsontofile) as data_file:
                        data = json.load(data_file)
                    if "test_cases" in data:
                        for i in data["test_cases"]:
                            for x in i:
                                durationcell += RX + ': ' + str(i[x]['duration']) + ' '
                    else:
                        jsontosavecell = "Malformed JSON"
                except IOError:
                     print "there is not reachability to ", jsonurl
                     jsontosavecell = "can't reach url"

            resultssheet[celdaU] = durationcell
            resultssheet[celdaS] = jsontosavecell

        elif re.match( regexx4url, jsonurlcell):
            print jsonurlcell
            jsontosave = re.match( regexx4url, jsonurlcell).group(1) + re.match( regexx4url, jsonurlcell).group(2) +'_'+ re.match( regexx4url, jsonurlcell).group(3) +'.json'

            dirjsontofile = jsondirectory + '/'+ jsontosave
            jsonurl = jsonurlcell

            try:
                fp = open(dirjsontofile, 'wb')
                req = urlopen(jsonurl)
                for line in req:
                    fp.write(line)
                fp.close()
                with open(dirjsontofile) as data_file:
                    data = json.load(data_file)
                if "test_cases" in data:
                    for i in data["test_cases"]:
                        for x in i:
                            resultssheet[celdaU] = i[x]['duration']
                            resultssheet[celdaS] = jsontosave
                else:

                    resultssheet[celdaS] = "Malformed JSON"
            except IOError:
                 print "there is not reachability to ",jsonurl
                 resultssheet[celdaS] = "can't reach url"
        elif jsonurl != '':
            print "no matching expression for: ", jsonurlcell

    a += 1
    print "row", a

outputfile = trkrfl[0] + timestr + ".xlsx"
wb.save(outputfile)
print "output file save at", outputfile

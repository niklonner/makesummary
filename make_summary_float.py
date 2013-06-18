# -*- coding: utf-8 -*-

import re
import sys
import os
import subprocess
import locale

def run(files):
#    locale.setlocale(locale.LC_ALL,'sv_SE.utf8');
    persons = []
    for _fileTmp in files:
        _file = open(os.path.join(sys.argv[1], _fileTmp),'r')    
        #print _file
        persons.append(makePerson(_file.readlines()))
        _file.close()
    print unlines(makeWorkbook(persons))

def makePerson(lines):
    _iter = lines.__iter__()
    person = {}
    person['name'] = re.split(r"\s*\n", _iter.next())[0]
    person['type'] = re.split(r"\s*\n", _iter.next())[0]
    person['startsum'] = float(_iter.next().split()[0].replace(',','.'))
    person['events'] = []
    for event in _iter:
        if re.split(r"\n", event)[0] == "":
            continue
        _event = {}
        eventIter = re.split(r"\s*;\s*", re.split(r"\n", event)[0]).__iter__()
        _event['date'] = eventIter.next()
        _event['description'] = eventIter.next()
        _event['amount'] = float(eventIter.next().replace(',','.'))
        person['events'].append(_event)
    return person
    
def makeSummary(persons):
    persons.sort(lambda x,y: cmp(x['type'].lower(),y['type'].lower()))
    lines = []
    lines.append("<ss:Worksheet ss:Name=\"Summering\">")
    lines.append("<ss:Table>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append(makeRow("Namn","Kvar att betala under året"))
    totalDebt = 0
    for person in persons:
        _currentDebt = -currentDebt(person)
        totalDebt += _currentDebt
        lines.append(makeRow(person['name'], _currentDebt, person['type']))
    lines.append(makeRow("Total utestående summa", totalDebt))
    lines.append("</ss:Table>")
    lines.append("</ss:Worksheet>")
    return lines

def makeWorkbook(persons):
    workbook = []
    workbook.append("<?xml version=\"1.0\"?>")
    workbook.append("<ss:Workbook xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">")
    workbook += makeSummary(persons)
    persons.sort(lambda x,y: cmp(x['name'].lower(),y['name'].lower()))
    for person in persons:
        workbook += makeStrPerson(person)
    workbook.append("</ss:Workbook>")
    return workbook

def flatFeeFromType(_type):
    _type = _type[:2].lower()
    if _type == "sj":               #
        return -1                   # Replace cases and values with
    elif _type == "fu":             # payment plans and their 
        return -1;                  # respective costs.
    elif _type == "ej":             #
        return -1;                  #
    elif _type == "ju":             #
        return -1;                  #
    return None

def currentDebt(person):
    debt = -flatFeeFromType(person['type'])
    debt += person['startsum']
    for event in person['events']:
        debt += event['amount']
    return debt

def makeStrPerson(person):
    lines = []
    lines.append("<ss:Worksheet ss:Name=\"" + person['name'] + "\">")
    lines.append("<ss:Table>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append("<ss:Column ss:Width=\"200\"/>")
    lines.append(makeRow(person['name'], "", person['type']))
    lines.append(makeRow("Ingående summa", "", person['startsum']))
    for event in person['events']:
        lines.append(makeRow(event['date'],event['description'],event['amount']))
    lines.append(makeRow("Nuvarande status", "", currentDebt(person)))
    lines.append("</ss:Table>")
    lines.append("</ss:Worksheet>")
    return lines

def makeRow(*cells):
    return reduce(lambda x,y: x + "<ss:Cell><ss:Data ss:Type=\"String\">" \
                            + str(y) + "</ss:Data></ss:Cell>", cells, "<ss:Row>") \
                            + "</ss:Row>"

def unlines(lines):
    return reduce(lambda x,y: x + "\n" + y, lines)

if __name__ == '__main__':
    # usage make_summary directory-of-data-files
    # remove tilde files in directory
    subprocess.call(["rm " + os.path.join(sys.argv[1], "*~")],shell=True)
    run(os.listdir(sys.argv[1]))
#    run(["niklas_lonnerfors.txt"])    

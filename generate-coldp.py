# -*- coding: utf-8 -*-

import csv, re
from openpyxl import Workbook, load_workbook
from collections import namedtuple
from unidecode import unidecode

xlsName   = 'ALF_Animalia_K-F_2019 Final_2019xii16 2.xlsx'
sheetNames= ['Animalia-Kingdom to Family']
notesCol  = 'T'
ranks     = ['kingdom','subkingdom','infrakingdom', 'superphylum','phylum','subphylum','infraphylum', 'superclass','class','subclass','infraclass', 'superorder', 'order', 'suborder', 'infraorder', 'series', 'subseries', 'superfamily', 'family']

#xlsName   = 'ALF_Non-Animalia_K-F_2019 Final_2019xii19.xlsx'
#sheetNames= ['Archaea-Bacteria', 'Protozoa', 'Chromista', 'Fungi', 'Plantae']
#notesCol  = 'U'
#ranks     = ['kingdom','subkingdom','infrakingdom', 'superphylum','phylum','subphylum','infraphylum','parvphylum', 'superclass','class','subclass','infraclass', 'superorder', 'order', 'suborder', 'infraorder', 'series', 'subseries', 'superfamily', 'family']

firstRow  = 2
firstCol  = 'A'
outFile   = 'NameUsage.csv'


Taxon = namedtuple('Taxon', 'id col name notes')
refMatcher = re.compile('\\b([A-Z][a-z]+) *(?:et al.?)?[, ]*(\\d{4})\\b')
synMatcher = re.compile('^(.+) *\[= *(.+) *] *')

parents = []    
sheet   = None
IDprefix  = ''

def read(row):
    note = sheet[notesCol + str(row+1)].value
    for col in range(1, len(ranks)+1):
        xxx = chr(ord(firstCol) + col - 1) + str(row+1)
        val = sheet[xxx].value
        #print(f'{xxx} => {val}')
        if val:
            return Taxon(str(row), col, val, note)
    if note:
        return Taxon(str(row), col, None, note)
    return None

def writeUsage(out, row, t, parentID, status):
    refID  = ""
    refIDs = []
    if t.notes:
        ascii = unidecode(t.notes)
        # might be several references separated by semicolon
        for note in ascii.split(';'):
            m = refMatcher.search(note)
            if m:
                refIDs.append(m.group(1).lower()+m.group(2))
    if refIDs:
        refID = '|'.join(refIDs)
    out.write("%s:%s,%s,%s,%s,%s,\"%s\",\"%s\",\"%s\"\n" % (IDprefix, t.id, parentID, row, status, ranks[t.col-1], t.name.replace('"', '""'), refID, (t.notes or "").replace('"', '""')))



wb = load_workbook(filename = xlsName)
with open(outFile, 'w', newline='') as out:
    out.write("ID,parentID,ordinal,status,rank,scientificName,referenceID,remarks\n")
    for sheetName in sheetNames:
        sheet = wb[sheetName]
        IDprefix = sheetName[0:2]
        row = 1 
        t = read(row)
        while (t):
            #print(t)
            if (t.name):
                while(parents and parents[-1].col >= t.col):
                    parents.pop()
                pid = IDprefix + ':' + parents[-1].id if parents else None
                m = synMatcher.search(t.name)
                if m:
                    s = Taxon('s'+str(row), t.col, m.group(2), None)
                    t = Taxon(t.id, t.col, m.group(1), t.notes)
                    writeUsage(out, row, s, t.id, 'synonym')
                writeUsage(out, row, t, pid or '', 'accepted')
                parents.append(t)
            row = row + 1
            t = read(row)

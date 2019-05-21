import xml.etree.ElementTree as ET
import sys
from pathlib import Path
from datetime import datetime, timedelta

def convertiOrario(dataOra):
    stringa = dataOra.split('T')
    data = stringa[0]
    ora = stringa[1]
    strOra = ora.split('-')
    ora = strOra[0][:-4]

    try:
        dataCorrente = datetime.strptime(data + ' ' + ora, '%Y-%m-%d %H:%M:%S')
        dataCorrente = dataCorrente + timedelta(hours=9)
    except:
        dataCorrente = 'ERR'

    return dataCorrente


if __name__ == '__main__':
    if len(sys.argv) != 2:
        print('Uso: gmailXMLParser.py nomeFileXML.xml')
    else:
        fXML = Path(str(sys.argv[1]))

        if fXML.is_file() and str(sys.argv[1]).upper().endswith('XML'):
            xmlfile = str(sys.argv[1])
            totMail = 0
            tree = ET.parse(xmlfile)
            root = tree.getroot()
            fCSV = open(str(sys.argv[1])+'.csv', 'w')
            fCSV.write('FROM;TO;CC;BCC;SUBJECT;DATA\n')
            for child in root:
                for documents in child:
                    for doc in documents:
                        try:
                            totMail += 1
                            mailID = doc.attrib['DocID']
                            frm = ''
                            to = ''
                            cc = ''
                            bcc = ''
                            subject = ''
                            labels = ''
                            dateSent = ''
                            dateReceived = ''
                            # INIZIO MAIL
                            for tags in doc:
                                for tag in tags:
                                    try:
                                        # print(tag.attrib['TagName'] + ': ' +  tag.attrib['TagValue'])
                                        if tag.attrib['TagName'] == '#From':
                                            frm = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#To':
                                            to = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#CC':
                                            cc = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#BCC':
                                            bcc = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#Subject':
                                            subject = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == 'Labels':
                                            labels = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#DateSent':
                                            dateSent = tag.attrib['TagValue']
                                        elif tag.attrib['TagName'] == '#DateReceived':
                                            dateReceived = tag.attrib['TagValue']
                                    except:
                                        pass
                            dateSent = convertiOrario(dateSent)
                            # print(type(dateSent))
                            # print(frm + ';' + to + ';' + cc + ';' + bcc + ';' + subject + ';' + str(dateSent) + '\n')
                            fCSV.write(frm + ';' + to + ';' + cc + ';' + bcc + ';' + subject + ';' + str(dateSent) + '\n')
                            # FINE MAIL
                        except:
                            pass
            # print('TOTALE EMAIL: ' + str(totMail))
            fCSV.close()
        else:
            print('File inesistente o tipo file non riconosciuto')
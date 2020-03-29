# -*- coding=gbk -*-
import glob
import sys

reload(sys)
sys.setdefaultencoding('UTF-8')
import csv
import re
import os
from xml.etree.ElementTree import iterparse
from HTMLParser import HTMLParser

importance_map = ['low', 'medium', 'high']
status_map = ['draft', 'to review', 'reviewing', 'rewritting', 'invalid', 'later versions', 'final']


class XML_CSV():
    def strip_tags(self, htmlStr):
        htmlStr = htmlStr.strip()
        htmlStr = htmlStr.strip("\n")
        # 删除style标签
        re_style = re.compile('<\s*style[^>]*>[^<]*<\s*/\s*style\s*>', re.I)
        htmlStr = re_style.sub('', htmlStr)
        result = []
        parser = HTMLParser()
        parser.handle_data = result.append
        htmlStr = parser.unescape(htmlStr)
        parser.feed(htmlStr)
        parser.close()
        return ''.join(result)

    def convert_xml2csv(self, csv_file, xmlfile):
        csvfile = open(csv_file, 'wb')
        spamwriter = csv.writer(csvfile, dialect='excel', delimiter=';')
        spamwriter.writerow(['TCID', 'CASE_NAME', 'IMPORTANCE', 'STATUS', 'SUMMARY', 'STEP', 'Result'])
        for (event, node) in iterparse(xmlfile, events=['end']):
            if node.tag == "testcase":
                case_list = ['', '', '', '', '', '', '']
                steps_list = ['', '', '', '', '', '', '']
                case_list[1] = node.attrib['name']
                for child in node:
                    if child.tag == "externalid":
                        text = re.sub('\n|<p>|</p>|\t', '', str(child.text))
                        # print self.strip_tags(text)
                        TCID = self.strip_tags(text)
                    elif child.tag == "summary":
                        text = re.sub('\n|<p>|</p>|\t', '', str(child.text))
                        # print self.strip_tags(text)
                        case_list[4] = self.strip_tags(text)
                    elif child.tag == "importance":
                        # text = re.sub('\n|<p>|</p>|\t', '', str(child.text))
                        # print self.strip_tags(text)
                        case_list[2] = importance_map[int(self.strip_tags(child.text)) - 1]
                    elif child.tag == "status":
                        # text = re.sub('\n|<p>|</p>|\t', '', str(child.text))
                        # print self.strip_tags(text)
                        case_list[3] = status_map[int(self.strip_tags(child.text)) - 1]
                        if "steps" not in [item.tag for item in node]:
                            case_list[0] = TCID
                            spamwriter.writerow(case_list)
                            break
                    elif child.tag == "steps":
                        if len(child) > 0:
                            for i in range(len(child)):
                                if i == 0:
                                    for n in range(len(child.getchildren()[i])):
                                        if child.getchildren()[i].getchildren()[n].text is not None:
                                            text = self.strip_tags(child.getchildren()[i].getchildren()[n].text).encode('UTF-8')
                                        else:
                                            text = ''
                                        # print text
                                        if child.getchildren()[i].getchildren()[n].tag == 'actions':
                                            case_list[5] = text
                                        elif child.getchildren()[i].getchildren()[n].tag == 'expectedresults':
                                            case_list[6] = text
                                    case_list[0] = TCID
                                    spamwriter.writerow(case_list)
                                else:
                                    for n in range(len(child.getchildren()[i])):
                                        if child.getchildren()[i].getchildren()[n].text is not None:
                                            text = self.strip_tags(child.getchildren()[i].getchildren()[n].text).encode('UTF-8')
                                        else:
                                            text = ''
                                        # print text
                                        if child.getchildren()[i].getchildren()[n].tag == 'actions':
                                            steps_list[5] = text
                                        elif child.getchildren()[i].getchildren()[n].tag == 'expectedresults':
                                            steps_list[6] = text
                                    steps_list[0] = TCID
                                    spamwriter.writerow(steps_list)
        csvfile.close()


if __name__ == "__main__":
    test = XML_CSV()
    # test.convert_xml2csv('output.csv', 'input.xml')
    for xml_file in glob.glob(os.curdir + '/*.xml'):
        test.convert_xml2csv(os.path.splitext(xml_file)[0] + '.csv', xml_file)

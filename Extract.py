# https://github.com/nmolivo/tesu_scraper/blob/master/Python_Blogs/01_extract_from_MSWord.ipynb
#specific to extracting information from word documents
import os
import zipfile
#other tools useful in extracting the information from our document
import re
#to pretty print our xml:
import xml.dom.minidom
#to read/write ics cal files
from ics import Calendar, Event
#use for formatting dates
from datetime import datetime

def extract(doc):
    #document will be the filetype zipfile.ZipFile
    document = zipfile.ZipFile(doc)
    # name = 'word/people.xml'
    uglyXml = xml.dom.minidom.parseString(document.read('word/document.xml')).toprettyxml(indent='  ')

    text_re = re.compile(r'>\n\s+([^<>\s].*?)\n\s+</', re.DOTALL)
    prettyXml = text_re.sub(r'>\g<1></', uglyXml)

    #print(prettyXml)

    # first to turn the xml content into a string:
    xml_content = document.read('word/document.xml')
    document.close()
    xml_str = str(xml_content)

    link_list = re.findall(r'(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', xml_str)
    link_list = link_list

    print(link_list)

    return link_list

def createIcs(doc):

    list_of_dates = extract(doc);

    for x in list_of_dates:
        try:
            date_object = datetime.strptime(x, "%m/%d/%y")

            #cutoff_year = 50

            #if (determine_year_length(date_object.year) == 2):
                #if (date_object.year > cutoff_year):
                  #  date_object = date_object.replace(year=date_object.year + 1900)
                #else:
                 #   date_object = date_object.replace(year=date_object.year + 2000)

            formatted_date = date_object.strftime("%Y-%m-%d")
        #formatted_date = formatted_date.split(" ")
        #formatted_date[-1] = formatted_date[-1][:4]
        #formatted_date = " ".join(formatted_date)

            print(formatted_date)
            c = Calendar()
            e = Event()
            e.name = input("Please insert name: ")
            e.begin =  f'{formatted_date} 00:00:00'
            c.events.add(e)
            c.events
            # {<Event 'My cool event' begin:2014-01-01 00:00:00 end:2014-01-01 00:00:01>}
            with open(f'{formatted_date}.ics', 'w') as f:
                f.writelines(c.serialize_iter())
                f.close()
            # And it's done !
            # iCalendar-formatted data is also available in a string
            print(c.serialize())
            # 'BEGIN:VCALENDAR\nPRODID:...
        except ValueError:
            print('ValueError Raised:', e)

def determine_year_length(year_str):
    year_str = str(year_str)
    year_length = len(year_str)

    return year_length


from odf.opendocument import load
from odf import text
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import collections

def Text(elem):
	total = ""
	tort = elem.text
	if tort: total += tort
	for child in elem:
		total += Text(child)
	return total

def convert_Title(p):
  text = '= ' + p
  return text

def convert_Heading(h, level):
  text = '=' * (level+1) + ' ' + h
  return text

def ListElement(elem):
  total = ''
  tort = elem.text
  text = []
  if tort: text += tort
  for child in elem:
    text.append(Text(child))
  return text

def convert_List(List, total):
  l = len(List)
  for i in range(0, l):
    text = []
    text = '* ' + List[i]
    total.append(text)
  return total

def TableElement(table):
  total = []
  for rows in table:
    for col in rows:
      total.append(Text(col))
  return total

def data_table(table, rows, cols, total):
  table_elements = TableElement(table)
  line = '|==='
  total.append(line)
  for i in range(rows):
    line = ''
    for j in range(cols):
      if table_elements[0] == '':
        line = line + '|' + '  '
      else:
        line = line + '|'+ table_elements[0] + ' ' 
      table_elements = table_elements[1:]
    total.append(line)
  line = '|==='
  total.append(line)
  return total

def write_in_adoc(d):
  file = open("convert.adoc", "w+")
  data = iter(d)
  for item in data:
    if item == '|===':
      file.write("%s\n" % item)
      next_item = next(data)
      while next_item != '|===':
        file.write("%s\n" % next_item)
        next_item = next(data)
      file.write("%s\n\n" % item)
    else:
      file.write("%s\n\n" % item)
  file.close()
  return

def odtToadoc(odtPath, total =[]):
  doc = load(odtPath)
  with ZipFile(odtPath, 'r') as odtArchive:
    try:
      with odtArchive.open(u'content.xml') as f:
        odtContent = f.read()
    except Exception as e:
      print("Could not find 'content.xml': {}".format(str(e)))
      return
    root = ET.fromstring(odtContent)
    for child in root.find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}body').find('{urn:oasis:names:tc:opendocument:xmlns:office:1.0}text'):
      if (child.tag == '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}h') and 'Heading' in str(child.attrib.values()):
        text = str(Text(child))
        if text == '':
          pass
        else:
          [level] = collections.deque(child.attrib.values(), maxlen=1)
          heading_adoc = convert_Heading(text, int(level))
          total.append(heading_adoc)
      elif (child.tag == '{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p'):
        if 'Title' in str(child.attrib.values()):
          text = str(Text(child))
          if text == '':
            pass
          else:
            Title_adoc = convert_Title(text)
            total.append(Title_adoc)
        elif ('Text' in str(child.attrib.values())):
          text = str(Text(child))
          if text == '':
            pass
          else:
            total.append(text)
        elif ('P' or 'Standard' or 'Normal' in str(child.attrib.values())):
          text = str(Text(child))
          if text == '':
            for x in child:
              if (x.tag == '{urn:oasis:names:tc:opendocument:xmlns:drawing:1.0}frame'):
                total.append('Здесь фото')
          else:
            total.append(text)
        else:
          print(child.attrib)
      elif ('list' in child.tag):
        text = ListElement(child)
        if text == '':
          pass
        else:
          total = convert_List(text, total)
      elif 'table' in child.tag:
        rows, cols, k = 0, 0, 0
        for x in child:
          if ('TableLine' in str(x.attrib)):
            rows +=1
          for y in x:
            k += 1
        cols = int(k / rows)
        total = data_table(child, rows, cols, total)
    return total

if __name__ == '__main__':
    total, data = [], []
    data = odtToText("111.odt", total)
    write_in_adoc(data)

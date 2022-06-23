import docx
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
import copy
import textwrap
import numpy as np

def data_table(doc):
  all_tables = doc.tables
  data_tables = {i:None for i in range(len(all_tables))}
  for i, table in enumerate(all_tables):
      data_tables[i] = [[] for _ in range(len(table.rows))]
      for j, row in enumerate(table.rows):
          for cell in row.cells:
              data_tables[i][j].append(cell.text)
  return data_tables

def iter_block_items(parent):
  if isinstance(parent, docx.document.Document): 
      parent_elm = parent.element.body
  else:
      raise ValueError("something's not right")
  for child in parent_elm.iterchildren():
      if isinstance(child, CT_P):
        yield docx.text.paragraph.Paragraph(child, parent)
      elif isinstance(child, CT_Tbl):
        yield docx.table.Table(child, parent)

def check_Title_style(k, doc):
  text = doc.paragraphs[k].text
  for run in doc.paragraphs[k].runs:
    if run.italic:
      text = '_' + text + '_'
    if run.bold:
      text  = '*' + text + '*'
    break
  text = '= ' + text
  return text

def check_Heading_style(k, doc, style):
  text = doc.paragraphs[k].text
  for run in doc.paragraphs[k].runs:
    if run.italic:
      text = '_' + text + '_'
    if run.bold:
      text  = '*' + text + '*'
    break
  s = int(style[-1]) + 1
  text = (s * '=') + ' ' + text
  return text  

def check_style_text(k, doc):
  text = copy.deepcopy(doc.paragraphs[k])
  for run in (text.runs):
    xmlstr = str(run.element.xml)
    if 'pic:pic' in run.element.xml:
      run.text = 'Photo'
      break  
    if run.bold:
      run.text = '*' + str(run.text) + '*'
    if run.italic:
      run.text = '_' + str(run.text) + '_'
    if run.underline:
       run.text = '[underline]#' + str(run.text) + '#'
    if run.font.superscript:
      run.text = '^' + str(run.text) + '^'
    if run.font.subscript:
      run.text = '~' + str(run.text) + '~'
  return text.text

def check_List_Paragraph_style(k, doc):
  list_item = copy.deepcopy(doc.paragraphs[k])
  list_lvl = doc.paragraphs[k]._p.pPr.numPr.ilvl.val
  for run in (list_item.runs):
    if run.bold:
      run.text = '*' + str(run.text) + '*'
    if run.italic:
      run.text = '_' + str(run.text) + '_'
    if run.underline:
       run.text = '[underline]#' + str(run.text) + '#'
    if run.font.superscript:
      run.text = '^' + str(run.text) + '^'
    if run.font.subscript:
      run.text = '~' + str(run.text) + '~'
  list_item = ((list_lvl + 1) * '*') + ' ' + str(list_item.text)
  return list_item

def write_in_asccidoc_file(file, text, k):
  file = open("convert.adoc", 'a+')
  if k == 0:
    file.write(str(text) + '\n\n')
    file.close()
  else:
    file.write('\n' + str(text) + '\n\n')
    file.close()
  return file

def write_tables(file, data_tables, n, k):
  table_ascii = np.array(data_tables[n])
  file = open("convert.adoc", 'a+')
  if (k == 0) and (n == 0):
    file.write('|===' + '\n')
    for i in range(0, len(table_ascii)):
      file.write('|' + ' |'.join(table_ascii[i]) + '\n')
    file.write('|===' + '\n\n')
    file.close()    
  else:
    file.write('\n' + '|===' + '\n')
    for i in range(0, len(table_ascii)):
      file.write('|' + ' |'.join(table_ascii[i]) + '\n')
    file.write('|===' + '\n\n')
    file.close()
  return

def convert(doc, k = 0, number_table = 0):
  file = open("convert.adoc", "w+")
  file.close()
  data_tables = data_table(doc)
  for block in iter_block_items(doc):
    if 'text' in str(block):
      style = str(block.style.name)
      if 'Title' in style:
        title = check_Title_style(k, doc)
        write_in_asccidoc_file(file, title, k)
      elif "Heading" in style:
        heading = check_Heading_style(k, doc, style)
        write_in_asccidoc_file(file, title, k)
      elif 'Normal' in style:
        txt = check_style_text(k, doc)
        if txt == '':
          pass
        else:
          split_txt = textwrap.wrap(txt, 70)
          file = open("convert.adoc", 'a+')
          file.write('\n'.join(split_txt) + '\n\n')
          file.close()
      elif 'List Paragraph' in style:
          List = check_List_Paragraph_style(k, doc)
          file = open("convert.adoc", 'a+')
          file.write(List + '\n')
          file.close()
      else:
          print('12312')
      k += 1
    elif 'table' in str(block):
      write_tables(file, data_tables, number_table, k)
      number_table += 1
  return

if __name__=='__main__':
  document = '3.docx'
  doc = Document(document)
  convert(doc)
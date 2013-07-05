#!/usr/bin/python

import gspread
import smtplib
import sys

from datetime import datetime
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

class TeacherData:
  """Class to represent data for a teacher
  """

  def __init__(self, data_row):
    #print data_row
    self.name = data_row[0].strip()
    self.email_id = data_row[1].strip()
    self.dates = []
    dates_data = data_row[2:]
    i = 0
    while i < len(dates_data):
      try:
        str_date = '%s %s' % (dates_data[i], dates_data[i+1])
        self.dates.append(datetime.strptime(str_date, '%m/%d/%Y %H:%M:%S'))
        i += 2
      except IndexError:
        print ('Looks like you did not entered dates'
               'properly for %s' % self.name)
        sys.exit(1)
      except:
        print 'Failed to parse dates for %s' % self.name
        sys.exit(1)

  def __str__(self):
    s = ('%s, %s, %s' %
         (self.name, self.email_id, ', '.join([str(d) for d in self.dates])))
    return s

  def __repr__(self):
    dates_str = [','.join(str(d).split()) for d in self.dates]
    s = '[%s,%s,%s]' % (self.name, self.email_id,
                        ','.join(dates_str))
    return 'TeacherData(%s)' % s

class AllTeachersData(list):
  """Class to represnt information regarding teacher's name, email and sub date,
  from a given spread sheet of a giver user.
  """

  def __init__(self, user_id, passwd, spreadsheet_name):
    gc = gspread.login(user_id, passwd)
    sp = gc.open(spreadsheet_name)
    ws = sp.sheet1

    row_num = 2
    while True:
      curr_row = ws.row_values(row_num)
      if not curr_row:
        break

      self.append(TeacherData(curr_row))
      row_num += 1

  def __str__(self):
    return '\n'.join([str(td) for td in self])

if __name__ == '__main__':
  atd = AllTeachersData('jiteshkhasgiwala', 'Uiyqwt82qoq!', 'Sub Schedule')
  print atd


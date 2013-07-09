#!/usr/bin/python

import copy
import gspread
import smtplib
import sys

from datetime import datetime
from datetime import timedelta
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email import Message

CLASS_SCHEDULES = {
  'Sun 10:00 AM': 'Hatha - Basics',
  'Sun 05:00 PM': 'Hatha - All Levels',
  'Sun 07:00 PM': 'Restorative Yoga',
  'Mon 09:00 AM': 'Yoga Gentle',
  'Mon 06:00 PM': 'Therapeutic Yoga',
  'Mon 07:15 PM': 'Hatha - Flow',
  'Tue 09:00 AM': 'Hatha - All Levels',
  'Tue 06:00 PM': 'Kundalini',
  'Tue 07:30 PM': 'Hatha - All Levels',
  'Wed 09:00 AM': 'Therapeutic Yoga',
  'Wed 06:00 PM': 'Yoga Gentle',
  'Wed 07:15 PM': 'Hatha - All Levels',
  'Thu 06:15 AM': 'Sunrise Yoga',
  'Thu 12:00 PM': 'Hatha - All Levels',
  'Thu 06:00 PM': 'Hatha - All Levels',
  'Thu 07:30 PM': 'Yoga Strech Kripalu',
  'Fri 09:30 AM': 'Hatha - All Levels',
  'Fri 06:30 PM': 'Therapeutic Yoga',
  'Sat 09:30 AM': 'Hatha - Flow',
  'Sat 11:30 AM': 'Yoga - Gentle',
  'Sat 05:00 PM': 'Hatha - All Levels',
  }

class TeacherData:
  """Class to represent data for a teacher
  """
  def __init__(self, data_row):
    """
    Args:
      data_row: List of all cells of a row from a given spreadsheet
    """
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
               'correctly for %s' % self.name)
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

  def GetDates(self):
    """Returns a copy of list of all dates
    """
    return copy.deepcopy(self.dates)

  def GetFirstName(self):
    """Returns the first name of the teacher.
    """
    return self.name.split()[0].strip()

class AllTeachersData(list):
  """Class to represnt information regarding teacher's name, email and sub date,
  from a given spread sheet of a giver user.
  """

  def __init__(self, user_id, passwd, spreadsheet_name):
    """
    Args:
      user_id: Gmail user id
      passwd: Gmail password
      spreadsheet_name: Name of the spreadsheet containing all the subs info
    """
    try:
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
    except gspread.exceptions.AuthenticationError:
      print 'Incorrect username or password'
      sys.exit(1)
    except gspread.exceptions.SpreadsheetNotFound:
      print '\'%s\' spreadsheet not found' % spreadsheet_name
      sys.exit(1)

  def __str__(self):
    return '\n'.join([str(td) for td in self])

class SendReminder:
  """Class to decide whom to send emails for what dates.
  The algorithm is to only send email, if the teacher is
  subbing in the following week.
  """
  @staticmethod
  def Send(all_teachers_data, user_name, password):
    """
    Method to send the reminder emails.

    Args:
      all_teachers_data: List object containing the data for all teachers
      user_name: Gmail user name
      password: Gmail password
    """
    for teacher_data in all_teachers_data:
      potential_dates = SendReminder._GetPotentialDates(teacher_data.GetDates())

      if potential_dates:
        SendReminder._SendEmail(teacher_data.GetFirstName(),
                                teacher_data.email_id, potential_dates,
                                user_name, password)

  @staticmethod
  def _GetPotentialDates(dates):
    """Return all those dates which are within a week from today.

    Args:
      dates: List of all the dates a teacher is subbing
    Returns:
      List of those dates which are within a week from today
    """
    potential_dates = []
    today = datetime.today()
    for a_date in dates:
      diff_days = (a_date - today).days
      if diff_days >= 0 and diff_days <= 7:
        potential_dates.append(a_date)

    return sorted(potential_dates)

  @staticmethod
  def _SendEmail(name, recipient_email, sub_dates, from_user, password):
    """
    Method to compose and send an email to a teacher.

    Args:
      name: Name of the teacher
      recipient_email: Emaiuls of the teacher
      sub_dates: List of dates this teacher will be subbing this week
      from_user: Gmail user
      password: Gmail password
    """
    recipients = [recipient_email, from_user + '@gmail.com']

    msg = MIMEMultipart()
    msg['From'] = from_user
    msg['To'] = recipient_email
    msg['Cc'] = from_user
    msg['Subject'] = ('Sub Reminder for the week of %s at SoniYoga'
                      % datetime.today().strftime('%dth %B'))

    datefmt_str = '%a, %dth %B at %I:%M %p'
    namefmt_str = '%a %I:%M %p'
    yoga_classes = []
    for d in sub_dates:
      class_date = d.strftime(datefmt_str)
      try:
        class_name = CLASS_SCHEDULES[d.strftime(namefmt_str)]
        yoga_classes.append(class_date + ' for ' + class_name)
      except KeyError:
        print ('Class time does not match from the spreadsheet for %s.\n'
               'Spreadsheet date and time is %s' % (name, class_date))
        sys.exit(1)

    message = ('Namaste %s\n\nThis is just a reminder about your subs for the '
               'coming week. You have signed to teach the following classes'
               '.\n\n%s\n\nOm,\nSoniYoga' % (name, '\n'.join(yoga_classes)))

    html_message = ("""
    <html>
      <head></head>
      <body>
        <p>
          Namaste %s,<br><br>
          This is just a reminder about your subs for the
          coming week. You have signed to teach the following
          classes.<br><br>%s<br><br>
          Om,<br>
          <a href="https://clients.mindbodyonline.com/ASP/home.asp?studioid=2535">SoniYoga</a>
        </p>
      </body>
    </html>
    """ % (name, '<br>'.join(yoga_classes)))

    msg.attach(MIMEText(html_message, 'html'))

    print 'Sending email to %s' % name

    mailServer = smtplib.SMTP('smtp.gmail.com', 587)
    mailServer.ehlo()
    mailServer.starttls()
    mailServer.ehlo()
    mailServer.login(from_user, password)
    mailServer.sendmail(from_user, recipients, msg.as_string())
    mailServer.close()


if __name__ == '__main__':
  user_name = raw_input('Gmail User Id: ')
  password = raw_input('Password: ')
  atd = AllTeachersData(user_name, password, 'Sub Schedule')
  SendReminder.Send(atd, user_name, password)

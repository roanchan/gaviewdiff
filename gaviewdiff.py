"""A simple example of how to access the Google Analytics API."""
from __future__ import print_function
import gspread

from apiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials

import httplib2
import os
from oauth2client import client
from oauth2client import file
from oauth2client import tools
from apiclient import discovery
from oauth2client.file import Storage

from sys import version_info


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

CLIENT_SECRET_FILE = 'C:/Users/Administrator/Desktop/automation/client_secret.json'
APPLICATION_NAME = 'Google Sheets API ViewDiff'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    #credentials = store.get()
    #if not credentials or credentials.invalid:
    flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, ['https://www.googleapis.com/auth/analytics.readonly','https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive'])
    flow.user_agent = APPLICATION_NAME
    if flags:
        credentials = tools.run_flow(flow, store, flags)
    else: # Needed only for compatibility with Python 2.6
        credentials = tools.run(flow, store)
    #print('Storing credentials to ' + credential_path)
    return credentials

def get_service(api_name, api_version, scope, key_file_location,
                service_account_email):
  """Get a service that communicates to a Google API.

  Args:
    api_name: The name of the api to connect to.
    api_version: The api version to connect to.
    scope: A list auth scopes to authorize for the application.
    key_file_location: The path to a valid service account p12 key file.
    service_account_email: The service account email address.

  Returns:
    A service that is connected to the specified API.
  """

  credentials = ServiceAccountCredentials.from_p12_keyfile(
    service_account_email, key_file_location, scopes=scope)

  http = credentials.authorize(httplib2.Http())

  # Build the service object.
  service = build(api_name, api_version, http=http)

  return service


def get_first_profile_id(service):
  # Use the Analytics service object to get the first profile id.

  # Get a list of all Google Analytics accounts for this user
  accounts = service.management().accounts().list().execute()

  if accounts.get('items'):
    # Get the first Google Analytics account.
    account = accounts.get('items')[0].get('id')

    # Get a list of all the properties for the first account.
    properties = service.management().webproperties().list(
        accountId=account).execute()

    if properties.get('items'):
      # Get the first property id.
      property = properties.get('items')[0].get('id')

      # Get a list of all views (profiles) for the first property.
      profiles = service.management().profiles().list(
          accountId=account,
          webPropertyId=property).execute()

      if profiles.get('items'):
        # return the first view (profile) id.
        return profiles.get('items')[0].get('id')

  return None


def get_results(service, profile_id):
  # Use the Analytics Service Object to query the Core Reporting API
  # for the number of sessions within the past seven days.
  return service.data().ga().get(
      ids='ga:' + profile_id,
      start_date='7daysAgo',
      end_date='today',
      metrics='ga:sessions').execute()
	  
def get_view_data(analytics, viewId, startDate, endDate, metricsParam, dimensionsParam):
  data_values = {}
  
  response = analytics.reports().batchGet(
      body={
        'reportRequests': [
        {
          'viewId': viewId,
          'dateRanges': [{'startDate': startDate, 'endDate': endDate}],
          'metrics': metricsParam,
          'dimensions': dimensionsParam
        }]
      }
  ).execute()
  
  for report in response.get('reports', []):
    columnHeader = report.get('columnHeader', {})
    dimensionHeaders = columnHeader.get('dimensions', [])
    metricHeaders = columnHeader.get('metricHeader', {}).get('metricHeaderEntries', [])
   
    for row in report.get('data', {}).get('rows', []):
      dimensions = row.get('dimensions', [])
      dateRangeValues = row.get('metrics', [])

      key_dimension = ""

      for header, dimension in zip(dimensionHeaders, dimensions):
        # print(header + ': ' + dimension)
        key_dimension = key_dimension + '::' + dimension

      data_values.update({key_dimension: []})
      for j in range(len(metricsParam)): # This is just to tell you how to create a list.
        data_values[key_dimension].append(0)
		
      for i, values in enumerate(dateRangeValues):
        # print('Date range: ' + str(i))
		#j = 0
        for metricHeader, value in zip(metricHeaders, values.get('values')):
          print(metricHeader.get('name') + ': ' + value)
          data_values[key_dimension][i] = value
          print(data_values[key_dimension][i])
		  #j = j + 1
	  
		  
  return data_values

def write_to_sheets(sh, title, old_values, new_values, metricsParam, dimensionsParam):
  worksheet = sh.add_worksheet(title=title, rows="100", cols="20")
  
  worksheet.update_cell(1, 2, 'New View')
  worksheet.update_cell(1, 2+len(metricsParam), 'Old View')
  worksheet.update_cell(1, 2+(2*len(metricsParam)), 'New View/Old View - 100%')
  
  start_column_new = 2;
  start_column_old = 2+len(metricsParam);
  start_column_diff = 2+(2*len(metricsParam))
  
  for i in range(0, len(metricsParam)):
    worksheet.update_cell(2, i+start_column_new, metricsParam[i]['expression'])
    worksheet.update_cell(2, i+start_column_old, metricsParam[i]['expression'])
    worksheet.update_cell(2, i+start_column_diff, 'Diff '+metricsParam[i]['expression'])
  
  worksheet.update_acell('A2', 'Total')
  
  dimension_label = ''
  for i in range(0, len(dimensionsParam)):
    dimension_label = dimension_label + '::' + dimensionsParam[i]['name']
	
  worksheet.update_cell(5, 1, 'Old View')
  worksheet.update_cell(6, 1, dimension_label)
  worksheet.update_cell(5, start_column_new, 'New View')
  worksheet.update_cell(5, start_column_old, 'Old View')
  worksheet.update_cell(5, start_column_diff, 'New View/Old View - 100%')
  
  for i in range(0, len(metricsParam)):
   worksheet.update_cell(6, start_column_new+i, metricsParam[i]['expression'])
   worksheet.update_cell(6, start_column_old+i, metricsParam[i]['expression'])
   worksheet.update_cell(6, start_column_diff+i, 'Diff '+ metricsParam[i]['expression'])
  
  cell_row = 7
  total_old_values = []
  for i in range(len(metricsParam)): # This is just to tell you how to create a list.
    total_old_values.append(0)
  for key in old_values:
    worksheet.update_cell(cell_row, 1, key)
    for i in range(0, len(metricsParam)):
      worksheet.update_cell(cell_row, i+start_column_old, old_values[key][i])
      if key in new_values:
        worksheet.update_cell(cell_row, i+start_column_new, new_values[key][i])
        worksheet.update_cell(cell_row, i+start_column_diff, str((float(new_values[key][i])/float(old_values[key][i])-1)*100)+'%')
      else: 
        worksheet.update_cell(cell_row, i+start_column_new, 0)
        worksheet.update_cell(cell_row, i+start_column_diff, str((0/float(old_values[key][i])-1)*100)+'%') 
      total_old_values[i] = total_old_values[i] + float(old_values[key][i])  
    cell_row = cell_row + 1

  start_column_new_header = 3+(2*len(metricsParam))+len(metricsParam)
  start_column_new = start_column_new_header+1;
  start_column_old = start_column_new_header+1+len(metricsParam);
  start_column_diff = start_column_new_header+1+(2*len(metricsParam))
  
  worksheet.update_cell(5, start_column_new_header, 'New View')
  worksheet.update_cell(6, start_column_new_header, dimension_label)
  worksheet.update_cell(5, start_column_new, 'New View')
  worksheet.update_cell(5, start_column_old, 'Old View')
  worksheet.update_cell(5, start_column_diff, 'New View/Old View - 100%')
  
  for i in range(0, len(metricsParam)):
    worksheet.update_cell(6, start_column_new+i, metricsParam[i]['expression'])
    worksheet.update_cell(6, start_column_old+i, metricsParam[i]['expression'])
    worksheet.update_cell(6, start_column_diff+i, 'Diff '+ metricsParam[i]['expression'])
	
  cell_row = 7
  total_new_values = []
  for i in range(len(metricsParam)): # This is just to tell you how to create a list.
    total_new_values.append(0)
  for key in new_values:
    worksheet.update_cell(cell_row, start_column_new_header, key)
    for i in range(0, len(metricsParam)):
      worksheet.update_cell(cell_row, i+start_column_new, new_values[key][i])
      if key in new_values:
        worksheet.update_cell(cell_row, i+start_column_old, old_values[key][i])
        worksheet.update_cell(cell_row, i+start_column_diff, str((float(new_values[key][i])/float(old_values[key][i])-1)*100)+'%')
      else: 
        worksheet.update_cell(cell_row, i+start_column_old, 0)
        worksheet.update_cell(cell_row, i+start_column_diff, '100%') 
      total_new_values[i] = total_new_values[i] + float(new_values[key][i])  
    cell_row = cell_row + 1
	
  start_column_new = 2;
  start_column_old = 2+len(metricsParam);
  start_column_diff = 2+(2*len(metricsParam))
	
  for i in range(0, len(metricsParam)):	
    worksheet.update_cell(3, i+start_column_new, total_new_values)
    worksheet.update_cell(3, i+start_column_old, total_old_values)
    worksheet.update_cell(3, i+start_column_diff, str((total_new_values[i]/total_old_values[i]-1)*100)+'%')

def record_device_category(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:sessions'}]
  dimensionsParam = [{'name': 'ga:deviceCategory'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Device Category"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_country(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:sessions'}]
  dimensionsParam = [{'name': 'ga:country'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Country"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_traffic(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:sessions'}]
  dimensionsParam = [{'name': 'ga:source'},{'name': 'ga:medium'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Traffic Sources"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_hostname(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:sessions'}]
  dimensionsParam = [{'name': 'ga:hostname'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Hostnames"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_page(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:pageviews'}]
  dimensionsParam = [{'name': 'ga:pagePath'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Pages"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_event(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:totalEvents'},{'expression': 'ga:eventValue'}]
  dimensionsParam = [{'name': 'ga:eventCategory'},{'name': 'ga:eventAction'},{'name': 'ga:eventLabel'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Events"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_transaction(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:transactions'},{'expression': 'ga:transactionRevenue'}]
  dimensionsParam = [{'name': 'ga:transactionId'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Transactions"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)
  
def record_product(analytics, sh, start_date, end_date, old_view_id, new_view_id):
  
  metricsParam = [{'expression': 'ga:itemQuantity'},{'expression': 'ga:itemRevenue'}]
  dimensionsParam = [{'name': 'ga:productSku'},{'name': 'ga:productName'},{'name': 'ga:productCategory'}]
  
  old_sessions = get_view_data(analytics, old_view_id, start_date, end_date, metricsParam, dimensionsParam)
  new_sessions = get_view_data(analytics, new_view_id, start_date, end_date, metricsParam, dimensionsParam)
  
  title = "Diff Products"
  write_to_sheets(sh, title, old_sessions, new_sessions, metricsParam, dimensionsParam)

def main():

  py3 = version_info[0] > 2 # creates boolean value for test that Python major version > 2

  if py3:
    start_date = input("Please enter Start Date(YYYY-MM-DD): ")
    end_date = input("Please enter End Date(YYYY-MM-DD): ")
    old_view_id = input("Please enter Old View ID: ")
    new_view_id = input("Please enter New View ID: ")
  else:
    start_date = raw_input("Please enter Start Date(YYYY-MM-DD): ")
    end_date = raw_input("Please enter End Date(YYYY-MM-DD): ")
    old_view_id = raw_input("Please enter Old View ID: ")
    new_view_id = raw_input("Please enter New View ID: ")
  
  credentials = get_credentials()
  
  http = credentials.authorize(httplib2.Http())
  analytics = build('analytics', 'v4', http=http)

  gc = gspread.authorize(credentials)
  sh = gc.create('GAViewDiff')
  
  record_device_category(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_country(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_traffic(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_hostname(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_page(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_event(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_transaction(analytics, sh, start_date, end_date, old_view_id, new_view_id)
  record_product(analytics, sh, start_date, end_date, old_view_id, new_view_id)


if __name__ == '__main__':
  main()
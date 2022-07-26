
'''
_______________________________________________________________________________
 LJB
_______________________________________________________________________________

   Program:     asset-workorder-updates-ljb.py
   Purpose:     1. Grab the Work Order feature layers 
                and build the Asset, Work Order dictionaries. 
                2. For both the OP and PE assets:
                    2a. If a next-inspection date exists in the Asset, a related Work 
                    Order does not exist, and the due date is within 40/375 
                    days (operator/PE):
                        - Create a new work order, with attachment
                        - Add the inspection to the Upcoming Excel worksheet
                    2b. If a next-inspection date exists in the Asset, a related Work 
                    Order does exist, and the inspection is complete:
                        - Add the inspection to the Completed Excel worksheet
                        - Update the Last and Next inspection date fields in 
                        the Assets
                        - Calculate the next due date. If next due date is 
                            within 40/375 days (operator/PE):
                            - Create a new work order, with attachment
                            - Add the inspection to the Upcoming Excel 
                            worksheet
                            - Update the Last and Next inspection date fields 
                            in the Assets
                    2c. If a next-inspection date exists in the Asset, a related Work 
                    Order does exist, and the inspection is not complete:
                        - Compare the due date to the current date and determine 
                        if the inspection is overdue. 
                        - If overdue (due date < current date):
                            - Calculate the days overdue
                            - Add the inspection to the Overdue Excel worksheet
                            - Update Status to 'Overdue' in the Work Orders
                        - If not overdue (due date > current date):
                            - Add the inspection to the Upcoming Excel worksheet
                3. Populate the List feature layer with all work orders, 
                occurring at the same location.
                4. Create the Excel workbook.
                5. Populate the Excel workbook with upcoming, overdue, and 
                completed inspections for Operator and PE.
                6. Send email with Excel workbook attached to client and LJB.
_______________________________________________________________________________
   History:     GTG     06/2021    Created
                GTG     12/2021    Update process... LJB migrated away 
                                    from Workforce to a custom feature
                                    service and applications.
                GTG     07/2022    Updated process. See above for updated summary.
                                    Used permanent work order 
                                    points and distinct asset feature classes.
_______________________________________________________________________________
'''
#%%
from arcgis.gis import GIS

from datetime import datetime, timedelta
import json
import logging
import os
from os import path
import shutil
import tempfile

from xlsxwriter import Workbook

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders


def printLog(words):
    '''Prints and logs progress/errors/exceptions'''

    logging.info(words)
    print(words)

def sendemail(eto, efrom, subject, message, email_cc, att = ''):
    '''Sends email to client and to LJB with the
    Excel workbook of assignment updates attached'''

    mail_from = efrom
    mail_body = message

    printLog('Building email structure...')
    mimemsg = MIMEMultipart()
    mimemsg['From']= mail_from
    mimemsg['To']= eto
    mimemsg['Subject']= subject
    mimemsg["Cc"] = email_cc  # Recommended for mass emails
    mimemsg.attach(MIMEText(mail_body, 'plain'))

    if att != '':
        printLog('Adding attachment...')
        filename = path.basename(att)
        attachment = open(att, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        mimemsg.attach(part)

    printLog('Connecting to SMTP host...')
    connection = smtplib.SMTP(host='10.0.7.5')
    # connection.set_debuglevel(True)
    printLog('Sending email...')
    connection.send_message(mimemsg)
    connection.quit()

def createWorkbook(output_dir):
    ''' Creates Excel workbook Completed, Upcoming
    and Overdue assignments. '''

    today = datetime.today().strftime('%Y%m%d')
    wb_path = output_dir + r'\WWSAssetUpdates_{}.xlsx'.format(today)

    workbook = Workbook(wb_path)
    printLog('Excel workbook created...')

    return(workbook, wb_path)

def addWorksheet(wb, worksheet_info):
    '''Adds worksheet and data to supplied Excel workbook'''

    for sheetname, ws_info in worksheet_info.items():

        values = ws_info[0]
        cols = ws_info[1]

        sheet = wb.add_worksheet(sheetname)

        for c in cols:
            sheet.write(c[0], c[1])

        row = 1
        for v in values.values():
            row += 1
            for i, val in enumerate(v):
                sheet.write('{0}{1}'.format(chr(ord('@')+(i+1)), str(row)), val)

        printLog('Worksheet {} was created and populated...'.format(sheetname))

    return(wb, sheet)

def buildQueryDictionaries(orgURL, username, password, itemID, \
                            wo_table_index, wo_index, crane_index, 
                            forktruck_index, eyewash_index, fire_index):
    
    '''Builds dictionaries to be used for asset and work order updates 
    in the updateServices module.Asset dictionary is populated if a next 
    due date value exists.
    Work Order dictionaries are populated if work order due date matches 
    match the due date in the assets'''

    gis = GIS(orgURL, username, password)
    printLog('Connected to Portal as {}'.format(gis.properties['user']['username']))

    sdi_item = gis.content.get(itemID)
    work_orders_tbl = sdi_item.tables[wo_table_index]
    work_orders_lyr = sdi_item.layers[wo_index]
    crane_lyr = sdi_item.layers[crane_index]
    forktruck_lyr = sdi_item.layers[forktruck_index]
    eyewash_lyr = sdi_item.layers[eyewash_index]
    fire_lyr = sdi_item.layers[fire_index]

    printLog('Building queried feature layers...')
    work_orders_tbl_feat= work_orders_tbl.query()
    work_orders_feat = work_orders_lyr.query()
    crane_feat = crane_lyr.query()
    forktruck_feat = forktruck_lyr.query()
    eyewash_feat = eyewash_lyr.query()
    fire_feat = fire_lyr.query()

    asset_lyrs = [crane_lyr, fire_lyr, eyewash_lyr, forktruck_lyr]
    asset_list = [crane_feat, fire_feat, eyewash_feat, forktruck_feat]
    asset_dict = {} 

    printLog('Building asset features dictionary...')
    for asset_feat in asset_list:
        for a in asset_feat.features:
            # only add assets to dictionary if a 'NextInspection' has been set and has an 'AssetID'
            asset_id_fieldname = [k for k,v in a.attributes.items() if "assetid" in k.lower()][0]
            if a.attributes['NextInspection'] != None:
                if a.attributes[asset_id_fieldname] != None:
                    asset_dict[a.attributes[asset_id_fieldname]] = [
                        a.attributes['EquipType'],
                        a.attributes['MeltShopArea'],
                        a.attributes['Building'],
                        a.attributes['InspectNotes'],
                        a.attributes['InspectName'],
                        a.attributes['NextInspection'],
                        a.attributes['LastInspect'],
                        a.attributes['InspectInterval'],
                        a.attributes['Clock'],
                        a.attributes['OBJECTID'],
                        a.attributes['GlobalID'],
                        a.geometry]
                else:
                    printLog('Asset with Global ID {} is missing an Asset ID!!'.format(a.attributes['GlobalID']))

    wo_table_dict = {} 
    printLog('Building work order table dictionary...')
    for o in work_orders_tbl_feat.features:
        wo_table_dict[o.attributes['GlobalID']] = [
            o.attributes['username'],
            o.attributes['AssignmentStatus'],
            o.attributes['AssignmentType'],
            o.attributes['AssignmentDueDate'],
            o.attributes['LastInspect'],
            o.attributes['RELAssetID'],
            o.attributes['created_date'],
            o.attributes['OBJECTID'],
            o.geometry]

    work_orders_dict = {} 
    printLog('Building work order features dictionary...')
    for o in work_orders_feat.features:
        if o.attributes['RELAssetID'] in asset_dict.keys():
            work_orders_dict[o.attributes['RELAssetID']] = [
                o.attributes['AssignmentType'],
                o.attributes['username'],
                o.attributes['AssignmentDueDate'],
                o.attributes['AssignmentStatus'],
                o.attributes['LastInspect'],
                o.attributes['created_date'],
                o.attributes['GlobalID'],
                o.attributes['OBJECTID'],
                o.geometry]

    return(asset_lyrs, asset_list, work_orders_lyr, work_orders_tbl, asset_dict, wo_table_dict, work_orders_dict)
#%%
def updateServices(asset_lyrs, asset_list, work_orders_lyr, work_orders_tbl, asset_dict, wo_table_dict, work_orders_dict, today):
    
    '''Creates new work orders, updates status of overdue work orders, 
    and updates the last and next inspection dates of OP and PE assets.
    Outputs 6 dictionaries that are used to populate an Excel Workbook.'''

    # Asset Types
    asset_name_dict = {
        'Crane': 0,
        'Extinguisher': 1,
        'Eyewash': 2,
        'Forklift': 3
    }

    # Inspection Intervals
    insp_interval_dict = {
        'Daily': 86400000,
        'Shift Start': 43200000,
        'Weekly': 604800000,
        'Monthly': 2678400000,
        'End of Month': 2419200000
    }

    wb_overdue = {}
    wb_upcoming = {}
    wb_completed = {}

    crane_overdue = {} 
    crane_upcoming = {} 
    crane_completed = {}
    fire_overdue = {} 
    fire_upcoming = {}
    fire_completed = {}
    eyewash_overdue = {} 
    eyewash_upcoming = {}
    eyewash_completed = {}
    forktruck_overdue = {}
    forktruck_upcoming = {} 
    forktruck_completed = {} 

    
    printLog('Looping all assets with a due date...')
    # ---------------------------------------------------------------------------------------------------------------------------
    # get unique asset IDs in work order table
    wo_table_dict_assetids = []
    for k, v in wo_table_dict.items():
        if v[5] not in wo_table_dict_assetids:
            wo_table_dict_assetids.append(v[5])

    for k, v in asset_dict.items():
        # identify variables 
        asset_id = k
        asset_type, area, building, notes, insp_name, next_insp, last_insp, insp_interval, insp_clock, obj_id, global_id, geom = v

        # Reverse lookup asset by 'EquipType'          
        asset_lyr = asset_lyrs[asset_name_dict[asset_type]]

        print('Asset ID {}'.format(str(k)))

        if next_insp != None:
            # due date is set

            # format date fields for Excel
            next_insp_date = datetime.fromtimestamp(next_insp/1000)
            next_insp_format = next_insp_date.strftime("%Y-%m-%d %H:%M:%S")
            last_insp_format = datetime.fromtimestamp(last_insp/1000).strftime("%Y-%m-%d %H:%M:%S") if last_insp != None else "No previous inspection"

            if asset_id not in work_orders_dict:
                # due date set and no work order found

                print('Due date set but no work order point exists')
                # if the due date within a month's time (~ 40 days), record as upcoming, else overdue
                print('Calculating days till due...')
                days_till_due = ((next_insp_date-today).days) + 1
                if days_till_due <= 60 and days_till_due >= 0:
                    # record work order as upcoming in Excel
                    print('Adding to upcoming dictionary...')
                    wb_upcoming[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, last_insp_format, insp_interval, asset_id]

                elif days_till_due < 0:
                    # record work order as overdue in Excel
                    days_overdue = abs(days_till_due)
                    print('Adding to overdue dictionary...')
                    wb_overdue[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, days_overdue, last_insp_format, insp_interval, asset_id] 
                                    
            elif asset_id in work_orders_dict:
                # due date set and a work order point exists

                if asset_id in wo_table_dict_assetids:
                    # due date set, work order point exists, and work order table row exists

                    # Grab most recent work order table row for asset
                    most_recent_wo_globalid = ''
                    latest_date = 0
                    for k, v in wo_table_dict.items():
                        if v[5] == asset_id:
                            if v[4] != None and v[4] > latest_date:
                                latest_date = v[4]
                                most_recent_wo_globalid = k
                            elif v[4] == None:
                                most_recent_wo_globalid = k
                    most_recent_wo = work_orders_tbl.query("GlobalID = '{}'".format(most_recent_wo_globalid)).features[0].attributes

                    printLog("Most recent work order row Asset ID: {0} has GlobalID: {1}".format(asset_id,most_recent_wo_globalid))

                    if last_insp != None and most_recent_wo['created_date'] <= last_insp:
                        # due date set, work order point exists, work order table row exists, and work order table row is outdated

                        print('Work order is older than existing work order point...')
                            
                        print('Calculating days till due...')
                        days_till_due = ((next_insp_date-today).days) + 1

                        if days_till_due <= 60 and days_till_due >= 0:
                            # if the due date within a month's time (~ 40 days), edit work order and add to upcoming in Excel

                            print('Adding to upcoming dictionary...')
                            wb_upcoming[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, last_insp_format, insp_interval, asset_id]
                            
                            # update AssignmentStatus as 'Assigned' in existing work order point
                            wo_edit = {'attributes':{}}
                            wo_edit['attributes']['AssignmentStatus'] = "Assigned"
                            wo_edit['attributes']['NextInspection'] = next_insp
                            wo_edit['attributes']['LastInspect'] = last_insp
                            wo_edit['attributes']['Clock'] = insp_clock
                            wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                            try:      
                                printLog('Editing work order point...Asset ID {}'.format(asset_id))
                                work_orders_lyr.edit_features(updates=[wo_edit])
                            except Exception as e:
                                printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                                print(e)

                        elif days_till_due < 0:
                            # if the due date within a month's time (~ 40 days), edit work order and add to upcoming in Excel

                            # record work order as overdue in Excel
                            days_overdue = abs(days_till_due)
                            print('Adding to overdue dictionary...')
                            wb_overdue[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, days_overdue, last_insp_format, insp_interval, asset_id] 
                                            
                            # update AssignmentStatus as 'Overdue' in existing work order point
                            print('Updating the status to "Overdue" in Work Order point...')
                            wo_edit = {'attributes':{}}
                            wo_edit['attributes']['AssignmentStatus'] = "Overdue"
                            wo_edit['attributes']['NextInspection'] = next_insp_date
                            wo_edit['attributes']['LastInspect'] = last_insp
                            wo_edit['attributes']['Clock'] = insp_clock
                            wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                            try:      
                                printLog('Editing work order point...Asset ID {}'.format(asset_id))
                                work_orders_lyr.edit_features(updates=[wo_edit])
                            except Exception as e:
                                printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                                print(e)

                    elif last_insp == None or most_recent_wo['created_date'] > last_insp:
                        # due date set, work order point exists, and a recent work order exists 
                        
                        # Calculate and format dates
                        next_due_date = datetime.fromtimestamp((most_recent_wo['created_date'] + insp_interval_dict[insp_interval])/1000) if insp_interval in insp_interval_dict else next_insp
                        if insp_interval == 'End of Month':
                            next_month = next_due_date.replace(day=28) + timedelta(days=4)
                            next_due_date = next_month - timedelta(days=next_month.day)
                        next_due_date_format = next_due_date.strftime("%Y-%m-%d %H:%M:%S")
                        completed_date = datetime.fromtimestamp(most_recent_wo['created_date']/1000)
                        completed_date_format = completed_date.strftime("%Y-%m-%d %H:%M:%S")

                        # Edit work order point as "Completed"
                        wo_edit = {'attributes':{}}
                        wo_excluded_fields = []
                        for k, v in most_recent_wo.items():
                            if k not in wo_excluded_fields:
                                wo_edit['attributes'][k] = v
                        wo_edit['attributes']['AssignmentStatus'] = "Completed"
                        wo_edit['attributes']['NextInspection'] = next_due_date
                        wo_edit['attributes']['LastInspect'] = most_recent_wo['created_date']
                        wo_edit['attributes']['Clock'] = most_recent_wo['Clock']
                        wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                        try:      
                            printLog('Editing work order point...Asset ID {}'.format(asset_id))
                            work_orders_lyr.edit_features(updates=[wo_edit])
                        except Exception as e:
                            printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                            print(e)

                        # Edit asset fields
                        asset_edit = {'attributes':{}}
                        asset_edit['attributes']['NextInspection'] = next_due_date
                        asset_edit['attributes']['LastInspect'] = most_recent_wo['created_date']
                        asset_edit['attributes']['InspectName'] = most_recent_wo['InspectName']
                        asset_edit['Clock'] = most_recent_wo['Clock']
                        asset_edit['attributes']['OBJECTID'] = obj_id
                        try:      
                            printLog('Editing asset...Asset ID {}'.format(asset_id))
                            asset_lyr.edit_features(updates=[asset_edit])
                        except Exception as e:
                            printLog('Failed to edit asset for Asset ID {}'.format(str(asset_id)))
                            print(e)

                        # record work order as complete in Excel
                        print('Adding to completed dictionary...')
                        wb_completed[asset_id] = [area, building, asset_type, notes, most_recent_wo["Clock"], completed_date_format, next_due_date_format, insp_interval, asset_id]

                        # add attachments from work order table to work order point
                        wo_tbl_oid = most_recent_wo['OBJECTID']
                        wo_oid = work_orders_dict[asset_id][7]
                        if len(work_orders_tbl.attachments.get_list(oid = wo_tbl_oid)) > 0:
                            with tempfile.TemporaryDirectory() as dirpath:
                                try:
                                    print('Downloading attachments...')
                                    paths = work_orders_tbl.attachments.download(oid = wo_tbl_oid, save_path = dirpath)
                                except Exception as e:
                                    printLog('Failed to download attachment from work order table with Asset ID {0}, Asset OID {1}'.format(str(asset_id), str(wo_tbl_oid)))
                                    printLog(e)
                                for path in paths:
                                    try:
                                        print('Adding attachments...')
                                        work_orders_lyr.attachments.add(oid = wo_oid, file_path = path)
                                    except Exception as e:
                                        printLog('Failed to upload attachments for work order point with Asset ID {0}, Work Order OID {1}'.format(str(asset_id), str(wo_oid)))
            
                elif asset_id not in wo_table_dict_assetids:
                    # due date set, work order point exists, and work order row doesn't exist

                        print('Work order not created for existing work order point...')
                            
                        print('Calculating days till due...')
                        days_till_due = ((next_insp_date-today).days) + 1

                        if days_till_due <= 60 and days_till_due >= 0:
                            # if the due date within a month's time (~ 40 days), edit work order and add to upcoming in Excel

                            print('Adding to upcoming dictionary...')
                            wb_upcoming[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, last_insp_format, insp_interval, asset_id]
                            
                            # update AssignmentStatus as 'Assigned' in existing work order point
                            wo_edit = {'attributes':{}}
                            wo_edit['attributes']['AssignmentStatus'] = "Assigned"
                            wo_edit['attributes']['NextInspection'] = next_insp_date
                            wo_edit['attributes']['LastInspect'] = last_insp
                            wo_edit['attributes']['Clock'] = insp_clock
                            wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                            try:      
                                printLog('Editing work order point...Asset ID {}'.format(asset_id))
                                work_orders_lyr.edit_features(updates=[wo_edit])
                            except Exception as e:
                                printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                                print(e)

                        elif days_till_due < 0:
                            # if the due date within a month's time (~ 40 days), edit work order and add to upcoming in Excel

                            # record work order as overdue in Excel
                            days_overdue = abs(days_till_due)
                            print('Adding to overdue dictionary...')
                            wb_overdue[asset_id] = [area, building, asset_type, notes, insp_clock, next_insp_format, days_overdue, last_insp_format, insp_interval, asset_id] 
                                            
                            # update AssignmentStatus as 'Overdue' in existing work order point
                            print('Updating the status to "Overdue" in Work Order point...')
                            wo_edit = {'attributes':{}}
                            wo_edit['attributes']['AssignmentStatus'] = "Overdue"
                            wo_edit['attributes']['NextInspection'] = next_insp_date
                            wo_edit['attributes']['LastInspect'] = last_insp
                            wo_edit['attributes']['Clock'] = insp_clock
                            wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                            try:      
                                printLog('Editing work order point...Asset ID {}'.format(asset_id))
                                work_orders_lyr.edit_features(updates=[wo_edit])
                            except Exception as e:
                                printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                                print(e)
                        
                        elif days_till_due > 60:
                            # due date over one month away
                            
                            # copy asset point attributes to work order point
                            print('Updating the status to "Overdue" in Work Order point...')
                            wo_edit = {'attributes':{}}
                            wo_edit['attributes']['NextInspection'] = next_insp_date
                            wo_edit['attributes']['LastInspect'] = last_insp
                            wo_edit['attributes']['Clock'] = insp_clock
                            wo_edit['attributes']['OBJECTID'] = work_orders_dict[asset_id][7]
                            try:      
                                printLog('Editing work order point...Asset ID {}'.format(asset_id))
                                work_orders_lyr.edit_features(updates=[wo_edit])
                            except Exception as e:
                                printLog('Failed to edit work order point for Asset ID {}'.format(str(asset_id)))
                                print(e)

        # Assign excel worksheets
        if asset_name_dict[asset_type] == 0:
            if asset_id in wb_overdue:
                crane_overdue[asset_id] = wb_overdue[asset_id]
            elif asset_id in wb_completed:
                crane_completed[asset_id] = wb_completed[asset_id]
            elif asset_id in wb_upcoming:
                crane_upcoming[asset_id] = wb_upcoming[asset_id]
        elif asset_name_dict[asset_type] == 1:
            if asset_id in wb_overdue:
                fire_overdue[asset_id] = wb_overdue[asset_id]
            elif asset_id in wb_completed:
                fire_completed[asset_id] = wb_completed[asset_id]
            elif asset_id in wb_upcoming:
                fire_upcoming[asset_id] = wb_upcoming[asset_id]
        elif asset_name_dict[asset_type] == 2:
            if asset_id in wb_overdue:
                eyewash_overdue[asset_id] = wb_overdue[asset_id]
            elif asset_id in wb_completed:
                eyewash_completed[asset_id] = wb_completed[asset_id]
            elif asset_id in wb_upcoming:
                eyewash_upcoming[asset_id] = wb_upcoming[asset_id]
        elif asset_name_dict[asset_type] == 3:
            if asset_id in wb_overdue:
                forktruck_overdue[asset_id] = wb_overdue[asset_id]
            elif asset_id in wb_completed:
                forktruck_completed[asset_id] = wb_completed[asset_id]
            elif asset_id in wb_upcoming:
                forktruck_upcoming[asset_id] = wb_upcoming[asset_id]


# ---------------------------------------------------------------------------------------------------------------------------
    printLog('{} inspections are upcoming...'.format(str(len(wb_upcoming))))
    printLog('{} inspections are complete...'.format(str(len(wb_completed))))
    printLog('{} inspections are overdue...'.format(str(len(wb_overdue))))
    
    return(wb_upcoming, wb_completed, wb_overdue, \
            crane_upcoming, crane_completed, crane_overdue,
            forktruck_upcoming, forktruck_completed, forktruck_overdue,
            eyewash_upcoming, eyewash_completed, eyewash_overdue,
            fire_upcoming, fire_completed, fire_overdue)

def moveToListService(orgURL, username, password, itemID, wo_index, list_index):

    gis = GIS(orgURL, username, password)
    printLog('Connected to Portal as {}'.format(gis.properties['user']['username']))

    printLog('Grabbing feature services...')
    sdi_item = gis.content.get(itemID)
    work_orders_lyr = sdi_item.layers[wo_index]
    list_lyr = sdi_item.layers[list_index]

    printLog('Querying feature layers...')
    work_orders_feat = work_orders_lyr.query()
    list_feat = list_lyr.query()

    printLog('Copying work orders...')
    copy_feat = [f for f in work_orders_feat]
    printLog('Constructing deletion list...')
    del_feat = [f.attributes['OBJECTID'] for f in list_feat]       

    print('Updating the XY of work orders before copying over...')
    for feat in copy_feat:
        feat.geometry['x'] = -9453433.93 
        feat.geometry['y'] = 5067414.83 

    if len(del_feat) > 0:
        print('Deleting all rows from list...')
        list_lyr.edit_features(deletes = del_feat)

    print('Appending updated work orders...')
    list_lyr.edit_features(adds = copy_feat)

def cleanUp(archive_fldr, wb_path):
    '''Moves workbook to archived folder. '''

    wb_name = path.basename(wb_path)

    printLog('Moving workbook to archived...')
    shutil.move(wb_path, archive_fldr + r'\{}'.format(wb_name))

    printLog('Successfully moved to "archived"!')

    archived_files = os.listdir(archive_fldr)
    if len(archived_files) > 8:

        printLog('Cleaning up archive folder...')

        oldestxlsx = min([archive_fldr + r'\{}'.format(f) for f in archived_files if f.endswith('.xlsx')], key = path.getctime)

        if oldestxlsx != archive_fldr + r'\{}'.format(wb_name):
            os.remove(oldestxlsx)
#%%
if __name__ == "__main__":

    # ------------------------------------------ maintain log file ------------------------------------------
    today = datetime.now()
    logfile = (path.abspath(path.join(path.dirname(__file__), '..', r'logs\asset-workorder-updates-ljb-LOG_{0}_{1}.txt'.format(today.month, today.year))))
    logging.basicConfig(filename=logfile,
                        level=logging.INFO,
                        format='%(levelname)s: %(asctime)s %(message)s',
                        datefmt='%m/%d/%Y %I:%M:%S')

    printLog("Starting run... \n")

    # ------------------------------------------ local inputs  ------------------------------------------
    # working path
    workingfldr = (path.abspath(path.join(path.dirname(__file__), '..', r'data')))
    # archive
    archfldr = (path.abspath(path.join(path.dirname(__file__), '..', r'archived')))

    #  ------------------------------------------ get AGOL creds ------------------------------------------
    ago_text = open(workingfldr + r'\portal-creds.json').read()
    json_ago = json.loads(ago_text)

    # URL to ArcGIS Online organization or ArcGIS Portal
    orgURL = json_ago['orgURL']
    # Username of an account in the org/portal that can access and edit all services listed below
    username = json_ago['username']
    # Password corresponding to the username provided above
    password = json_ago['password']
    # Item ID to WWS Assets and Work Order service
    itemID = json_ago["itemid"]
    #asset_list = json_ago["assetindex"]
    wo_table_index = json_ago["tableindex"]
    wo_index = json_ago["workorderindex"]
    list_index = json_ago["listindex"]
    crane_index = json_ago["craneindex"]
    forktruck_index = json_ago["forktruckindex"]
    eyewash_index = json_ago["eyewashindex"]
    fire_index = json_ago["fireindex"]

    # ------------------------------------------ get email creds ------------------------------------------
    email_text = open(workingfldr + r'\email-info.json').read()
    json_email = json.loads(email_text)

    email_recipient = json_email['email_to']
    email_from = json_email['email_from']
    email_cc = json_email['email_cc']
    email_subject = json_email['subject']
    email_message = json_email['message']

    # ------------------------------------------ excel workbook columns ------------------------------------

    # "Completed" column names
    wb_complete_cols = [['A1', 'Area'],['B1', 'Building'],['C1', 'Equipment Type'], ['D1', 'Notes'],['E1', 'Inspector Clock'],['F1', 'Completed Date'],['G1', 'Next Inspection Date'],['H1', 'Inspection Interval'],['I1', 'Asset ID']]

    # "Upcoming" column names
    wb_upcoming_cols = [['A1', 'Area'],['B1', 'Building'],['C1', 'Equipment Type'],['D1', 'Notes'],['E1', 'Inspector Clock'],['F1', 'Next Inspection Date'],['G1', 'Last Inspection Date'],['H1', 'Inspection Interval'],['I1', 'Asset ID']]

    # "Overdue" column names
    wb_overdue_cols = [['A1', 'Area'],['B1', 'Building'],['C1', 'Equipment Type'],['D1', 'Notes'],['E1', 'Inspector Clock'],['F1', 'Duedate'],['G1', 'Days Overdue'],['H1', 'Last Inspection Date'],['I1', 'Inspection Interval'],['J1', 'Asset ID']]

#%%
    # ------------------------------------------ execute functions ------------------------------------------

    try: 
        printLog('Grabbing services and sorting data...')
        asset_lyrs, asset_list, work_orders_lyr, work_orders_tbl, asset_dict, wo_table_dict, work_orders_dict = buildQueryDictionaries(orgURL, username, password, itemID, \
                                                            wo_table_index, wo_index, crane_index, 
                                                            forktruck_index, eyewash_index, fire_index)
        printLog('Updating assets and work orders...')
        wb_upcoming, wb_completed, wb_overdue, crane_upcom, crane_com, crane_overdue, forktruck_upcom, forktruck_com, forktruck_overdue, eyewash_upcom, eyewash_com, eyewash_overdue, fire_upcom, fire_com, fire_overdue = updateServices(asset_lyrs, asset_list, work_orders_lyr, work_orders_tbl, asset_dict, wo_table_dict, work_orders_dict, today)
        
        printLog('Moving work orders to list...')
        moveToListService(orgURL, username, password, itemID, wo_index, list_index)

        worksheet_dict = {'Fire Extinguisher Completed': [fire_com, wb_complete_cols],
                        'Eyewash Completed': [eyewash_com, wb_complete_cols],
                        'Crane Completed': [crane_com, wb_complete_cols],
                        'Fork Truck Completed': [forktruck_com, wb_complete_cols],
                        'Fire Extinguisher Upcoming': [fire_upcom, wb_upcoming_cols],
                        'Eyewash Upcoming': [eyewash_upcom, wb_upcoming_cols],
                        'Crane Upcoming': [crane_upcom, wb_upcoming_cols],
                        'Fork Truck Upcoming': [forktruck_upcom, wb_upcoming_cols],
                        'Fire Extinguisher Overdue': [fire_overdue, wb_overdue_cols],
                        'Eyewash Overdue': [eyewash_overdue, wb_overdue_cols],
                        'Fork Truck Overdue': [forktruck_overdue, wb_overdue_cols],
                        'Crane Overdue': [crane_overdue, wb_overdue_cols]}
        
        printLog('Creating Excel workbook...')
        wb, wb_path = createWorkbook(workingfldr)
        printLog('Populating worksheets...')
        addWorksheet(wb, worksheet_dict)
        wb.close()

        if today.weekday() == 0:
            printLog('Sending e-mail...')
            sendemail(email_recipient, email_from, email_subject, email_message, email_cc, att=wb_path)
        else:
            printLog('Today is not Monday, skipping email send...')

        printLog('Clean up, aisle 9...')
        cleanUp(archfldr, wb_path)

        printLog("Success! \n ------------------------------------ \n\n")

    except Exception: 
        logging.error("EXCEPTION OCCURRED", exc_info=True)

        error_to = 'user@domain.com'
        error_from = 'user@domain.com'
        error_sub = 'FAILED asset-workorder-updates-ljb.py'
        error_msg = 'asset-workorder-updates-ljb.py failed. Log file is attached. Please review and contact GTG if necessary (sstokes@geotg.com or jrogers@geotg.com)'
        printLog("Sending log to owners...")

        try:
            sendemail(error_to, error_from, error_sub, error_msg, '', att = logfile)
        except Exception:
            logging.error("Email exception occurred...", exc_info=True)
            
        printLog("Quitting! \n ------------------------------------ \n\n")  


'''
_______________________________________________________________________________
 LJB
_______________________________________________________________________________

   Program:     asset-workorder-updates-ljb.py
   Purpose:     1. Grab the WWS Asset and Work Order feature layers 
                and build the Asset, OP Work Order, and PE Work Order 
                dictionaries. 
                2. For both the OP and PE assets:
                    2a. If a next due date exists in the Asset, a related Work 
                    Order does not exist, and the due date is within 40/375 
                    days (operator/PE):
                        - Create a new work order, with attachment
                        - Add the inspection to the Upcoming Excel worksheet
                    2b. If a next due date exists in the Asset, a related Work 
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
                    2c. If a next due date exists in the Asset, a related Work 
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
                GTG     12/2021    Update process... see above for 
                                    full summary. LJB migrated away 
                                    from Workforce to a custom feature
                                    service and applications.
_______________________________________________________________________________
'''

from arcgis.gis import GIS

from datetime import datetime
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

        printLog('Adding worksheet {}...'.format(sheetname))
        sheet = wb.add_worksheet(sheetname)

        printLog('Populating columns...')
        for c in cols:
            sheet.write(c[0], c[1])

        printLog('Populating rows...')
        row = 1
        for v in values.values():
            row += 1
            for i, val in enumerate(v):
                sheet.write('{0}{1}'.format(chr(ord('@')+(i+1)), str(row)), val)

        printLog('Worksheet {} was created and populated...'.format(sheetname))

    return(wb, sheet)


def buildQueryDictionaries(portal, un, pw, item_id, asset_index, wo_index):
    '''Builds dictionaries to be used for asset and work order updates 
    in the updateServices module.Asset dictionary is populated if a next 
    due date value exists.
    Work Order dictionaries are populated if work order due date matches 
    match the due date in the assets'''

    gis = GIS(portal, un, pw)
    printLog('Connected to Portal as {}'.format(gis.properties['user']['username']))

    sdi_item = gis.content.get(item_id)
    assets_lyr = sdi_item.layers[asset_index]
    work_orders_lyr = sdi_item.layers[wo_index]

    printLog('Building queried feature layers...')
    assets_feat = assets_lyr.query()
    op_work_orders_feat = work_orders_lyr.query("AssignmentType IN ('WWS Operator Inspection', 'Crane Operator Inspection')")
    pe_work_orders_feat = work_orders_lyr.query("AssignmentType = 'PE Inspection'")

    asset_dict = {} #HazardID: [type-0, area-1, location-2, description-3, description notes-4, 
        #                       production accessible-5, OP frequency-6, Next OP inspection-7, 
        #                       Last OP inspection-8, PE frequency-9, next PE inspection-10, 
        #                       last PE inspection-11, oid-12, geometry-13]
    printLog('Building asset features dictionary...')
    for a in assets_feat.features:
        # only add assets to dictionary if a due date has been set for either OP or PE
        if a.attributes['NextOpInspect'] != None or a.attributes['NextPEInspect'] != None:
            if a.attributes['HazardID'] != None:
                asset_dict[a.attributes['HazardID']] = [
                    a.attributes['OperatorType'],
                    a.attributes['AssetArea'],
                    a.attributes['AssetLocation'],
                    a.attributes['AssetDescription'],
                    a.attributes['AssetDescriptionNotes'],
                    a.attributes['ProductionAccessible'],
                    a.attributes['OperatorFrequency'],
                    a.attributes['NextOpInspect'],
                    a.attributes['LastOpInspect'],
                    a.attributes['PEFrequency'],
                    a.attributes['NextPEInspect'],
                    a.attributes['InspectDate'],
                    a.attributes['OBJECTID'],
                    a.geometry]
            else:
                printLog('Asset ObjectID {} is mising an Asset ID (Hazard ID)!!'.format(a.attributes['OBJECTID']))

    op_work_order_dict = {} # HazardID: [type-0, complete date-1, assigned to-2, due date-3, status-4, oid-5, geom-6]
    printLog('Building OP work order features dictionary...')
    for o in op_work_orders_feat.features:
        # only add work orders to dictionary where AssignmentDueDate == NextOpInspect from WWS Assets
        # this avoids adding older completed work orders and overwriting the hazard id key
        hazard_id_wo = o.attributes['HazardID']
        if hazard_id_wo in asset_dict.keys():
            if o.attributes['AssignmentDueDate'] == asset_dict[hazard_id_wo][7]:
                op_work_order_dict[hazard_id_wo] = [
                    o.attributes['AssignmentType'],
                    o.attributes['CompleteDate'],
                    o.attributes['username'],
                    o.attributes['AssignmentDueDate'],
                    o.attributes['AssignmentStatus'],
                    o.attributes['OBJECTID'],
                    o.geometry]

    pe_work_order_dict = {} # HazardID: [type-0, complete date-1, assigned to-2, due date-3, status-4, oid-5, geom-6]
    printLog('Building PE work order features dictionary...')
    for p in pe_work_orders_feat.features:
        # only add work orders to dictionary where AssignmentDueDate == NextOpInspect from WWS Assets
        # this avoids adding older completed work orders and overwriting the hazard id key
        hazard_id_wo = p.attributes['HazardID']
        if hazard_id_wo in asset_dict.keys():
            if p.attributes['AssignmentDueDate'] == asset_dict[p.attributes['HazardID']][7]:
                pe_work_order_dict[p.attributes['HazardID']] = [
                    p.attributes['AssignmentType'],
                    p.attributes['CompleteDate'],
                    p.attributes['username'],
                    p.attributes['AssignmentDueDate'],
                    p.attributes['AssignmentStatus'],
                    p.attributes['OBJECTID'],
                    p.geometry]

    return(assets_lyr, work_orders_lyr, asset_dict, op_work_order_dict, pe_work_order_dict)


def updateServices(assets, work_orders, asset_dict, op_work_order_dict, pe_work_order_dict, today):
    '''Creates new work orders, updates status of overdue work orders, 
    and updates the last and next inspection dates of OP and PE assets.
    Outputs 6 dictionaries that are used to populate an Excel Workbook.'''

    # get completed, upcoming, and overdue tasks
    op_overdue = {} # hazard id : [orig due date, days overdue, area, location, desc, notes, freq, last inspect date, hazard id, prod acc, type]
    op_upcoming = {} # hazard id : [area, location, desc, notes, due date, freq, last inspect date, hazard id, prod acc, type]
    op_completed = {} # hazard id : [area, location, desc, notes, complete date, worker, hazard id, prod acc, type]

    pe_overdue = {} # hazard id : [orig due date, days overdue, area, location, desc, notes, freq, last inspect date, prod acc]
    pe_upcoming = {} # hazard id : [area, location, desc, notes, due date, freq, last inspect date, prod acc]
    pe_completed = {} # hazard id : [area, location, desc, notes, complete date, worker, prod acc]

    op_type_dict = {'Operator': ['WWS Operator Inspection', 'Operator Contractor'], 'Crane': ['Crane Operator Inspection', 'Crane Crew']}

    printLog('Looping all assets with a due date...')
    for k, v in asset_dict.items():

        # identify variables 
        hazard_id = k
        op_type, area, location, desc, notes, prod_acc, op_freq, op_next_insp, \
        op_last_insp, pe_freq, pe_next_insp, pe_last_insp, obj_id, geom = v

        print('Asset ID {}'.format(str(k)))

        # OPERATOR
        if op_next_insp != None:
            if hazard_id not in op_work_order_dict.keys():
                print('Due date set and no work order is scheduled')
                duedate = datetime.fromtimestamp(op_next_insp/1000)

                # if the due date within a month's time (~ 40 days), create workorder and record as upcoming
                print('Calculating days till due...')
                days_till_due = (today - duedate).days

                if days_till_due <= 40:
                    print('Creating new OP work order...')

                    # record work order as upcoming in Excel
                    duedate_format = duedate.strftime("%Y-%m-%d %H:%M:%S")
                    if op_last_insp != None:
                        last_insp_format = (datetime.fromtimestamp(op_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        last_insp_format = 'No previous inspection'
                    print('Adding to upcoming dictionary...')
                    op_upcoming[hazard_id] = [area, location, desc, notes, duedate_format, op_freq, last_insp_format,
                                                hazard_id, prod_acc, op_type]

                    # create work order as Assigned
                    add_dict = {'geometry': geom, 
                                'attributes': {'username': op_type_dict[op_type][1],
                                                'AssignmentStatus': 'Assigned',
                                                'AssignmentType': op_type_dict[op_type][0],
                                                'AssignmentDueDate': op_next_insp,
                                                'AssetArea': area, 
                                                'AssetLocation': location,
                                                'AssetDescription': desc,
                                                'AssetDescriptionNotes': notes,
                                                'HazardID': hazard_id}}

                    # create new work order with same geometry as asset
                    try:      
                        print('Adding new OP work order...Hazard ID {}'.format(hazard_id))
                        work_orders.edit_features(adds = [add_dict])
                    except Exception as e:
                        print('Failed to add new work order for Hazard ID {}'.format(str(hazard_id)))
                        print(e)
                    # adds_l.append(add_dict)

                    # add attachments
                    new_work_order = work_orders.query('HazardID = {}'.format(hazard_id)).features[0]
                    new_oid = new_work_order.attributes['OBJECTID']

                    if len(assets.attachments.get_list(oid = obj_id)) > 0:
                        with tempfile.TemporaryDirectory() as dirpath:
                            try:
                                print('Downloading attachments...')
                                paths = assets.attachments.download(oid = obj_id, save_path = dirpath)
                            except Exception as e:
                                print('Failed to download attachment from Hazard ID {0}, Asset OID {1}'.format(str(hazard_id), str(obj_id)))
                                print(e)
                            for path in paths:
                                try:
                                    print('Adding attachments...')
                                    work_orders.attachments.add(oid = new_oid, file_path = path)
                                except Exception as e:
                                    print('Failed to upload attachments for Hazard ID {0}, Work Order OID {1}'.format(str(hazard_id), str(new_oid)))

            else:
                # due date set and a work order exists
                if op_work_order_dict[hazard_id][4] == 'Completed': # and op_work_order_dict[hazard_id][1] != None
                    print('OP work order is complete...')

                    duedate = datetime.fromtimestamp(op_next_insp/1000)

                    # record work order as complete in Excel
                    print('Formatting complete date...')
                    complete_date_agol = op_work_order_dict[hazard_id][1]
                    complete_date = datetime.fromtimestamp(complete_date_agol/1000)
                    complete_date_format = complete_date.strftime("%Y-%m-%d %H:%M:%S")

                    worker = op_work_order_dict[hazard_id][2]

                    print('Adding to completed dictionary...')
                    op_completed[hazard_id] = [area, location, desc, notes, complete_date_format, worker, hazard_id, prod_acc, op_type]

                    print('Calculating the next due date and the days till due...')
                    next_due_date_agol = op_next_insp + (op_freq*86400000)
                    next_due_date = datetime.fromtimestamp(next_due_date_agol/1000)
                    days_till_due = (next_due_date - duedate).days

                    # if next inspection date is within 40 days 
                    # create new work order with Assigned status
                    # add to upcoming in Excel
                    if days_till_due <= 40:
                        print('Creating new OP work order...')

                        # add to upcoming in Excel
                        next_due_date_format = next_due_date.strftime("%Y-%m-%d %H:%M:%S")
                        if op_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(op_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to upcoming dictionary...')
                        op_upcoming[hazard_id] = [area, location, desc, notes, next_due_date_format, op_freq, last_insp_format,
                                                    hazard_id, prod_acc, op_type]

                        # create work order as Assigned
                        add_dict = {'geometry': geom, 
                                    'attributes': {'username': op_type_dict[op_type][1],
                                                'AssignmentStatus': 'Assigned',
                                                'AssignmentType': op_type_dict[op_type][0],
                                                'AssignmentDueDate': next_due_date_agol,
                                                'AssetArea': area, 
                                                'AssetLocation': location,
                                                'AssetDescription': desc,
                                                'AssetDescriptionNotes': notes,
                                                'HazardID': hazard_id}}

                        # create new work order with same geometry as asset
                        try:      
                            printLog('Adding new work order...Hazard ID {}'.format(hazard_id))
                            work_orders.edit_features(adds = [add_dict])
                        except Exception as e:
                            printLog('Failed to add new work order for Hazard ID {}'.format(str(hazard_id)))
                            print(e)

                        # add attachments
                        new_work_order = work_orders.query('HazardID = {0} AND AssignmentDueDate = {1}'.format(hazard_id, next_due_date_agol)).features[0]
                        new_oid = new_work_order.attributes['OBJECTID']

                        if len(assets.attachments.get_list(oid = obj_id)) > 0:
                            with tempfile.TemporaryDirectory() as dirpath:
                                try:
                                    print('Downloading attachments...')
                                    paths = assets.attachments.download(oid = obj_id, save_path = dirpath)
                                except Exception as e:
                                    printLog('Failed to download attachment from Hazard ID {0}, Asset OID {1}'.format(str(hazard_id), str(obj_id)))
                                    printLog(e)
                                for path in paths:
                                    try:
                                        print('Adding attachments...')
                                        work_orders.attachments.add(oid = new_oid, file_path = path)
                                    except Exception as e:
                                        printLog('Failed to upload attachments for Hazard ID {0}, Work Order OID {1}'.format(str(hazard_id), str(new_oid)))

                        # update LastOpInspect and NextOpInspect with CompleteDate and new AssignmentDueDate
                        print('Updating the last and next inspection dates in Assets...')
                        asset_edit = (assets.query('HazardID = {}'.format(hazard_id))).features[0]
                        asset_edit.attributes['LastOpInspect'] = complete_date_agol
                        asset_edit.attributes['NextOpInspect'] = next_due_date_agol
                        assets.edit_features(updates = [asset_edit])

                    else:
                        # work order is not within 40 days
                        complete_date_agol = op_work_order_dict[hazard_id][1]
                        next_due_date_agol = op_next_insp + (op_freq*86400000)

                        # update LastOpInspect and NextOpInspect with CompleteDate and new AssignmentDueDate
                        print('Upadting the last and next inspection dates in Assets...')
                        asset_edit = (assets.query('HazardID = {}'.format(hazard_id))).features[0]
                        asset_edit.attributes['LastOpInspect'] = complete_date_agol
                        asset_edit.attributes['NextOpInspect'] = next_due_date_agol
                        assets.edit_features(updates = [asset_edit])

                else:
                    wo_obj_id = op_work_order_dict[hazard_id][5]

                    orig_duedate = op_work_order_dict[hazard_id][3]
                    duedate = datetime.fromtimestamp(orig_duedate/1000)
                    duedate_format = duedate.strftime("%Y-%m-%d %H:%M:%S")

                    if duedate < today:
                        print('OP work order is overdue...')

                        print('Calculating the days overdue...')
                        days_overdue = (today - duedate).days

                        # record work order as overdue in Excel
                        if op_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(op_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to overdue dictionary...')
                        op_overdue[hazard_id] = [duedate_format, days_overdue, area, location, desc, notes, op_freq, 
                                                last_insp_format, hazard_id, prod_acc, op_type] 
                                        
                        # update AssignmentStatus as Overdue in existing work order
                        print('Updating the status to overdue in Work Orders...')
                        work_order_edit = (work_orders.query('OBJECTID = {}'.format(wo_obj_id))).features[0]
                        work_order_edit.attributes['AssignmentStatus'] = 'Overdue'
                        work_orders.edit_features(updates = [work_order_edit])

                    else:
                        print('OP work order is upcoming...')

                        # record work order as upcoming in Excel
                        if op_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(op_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to upcoming dictionary...')
                        op_upcoming[hazard_id] = [area, location, desc, notes, duedate_format, op_freq, last_insp_format, 
                                                    hazard_id, prod_acc, op_type]

    # ---------------------------------------------------------------------------------------------------------------------------

        # PE
        if pe_next_insp != None:
            if hazard_id not in pe_work_order_dict.keys():
                print('Due date set and no work order is scheduled')
                duedate = datetime.fromtimestamp(pe_next_insp/1000)

                # if the due date within a year's time (~ 375 days), create workorder and record as upcoming
                print('Calculating days till due...')

                days_till_due = (today - duedate).days
                if days_till_due <= 375:
                    print('Creating new PE work order...')

                    # record work order as upcoming in Excel
                    duedate_format = duedate.strftime("%Y-%m-%d %H:%M:%S")
                    if pe_last_insp != None:
                        last_insp_format = (datetime.fromtimestamp(pe_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                    else:
                        last_insp_format = 'No previous inspection'
                    print('Adding to upcoming dictionary...')
                    pe_upcoming[hazard_id] = [area, location, desc, notes, duedate_format, pe_freq, last_insp_format, 
                                                hazard_id, prod_acc]

                    # create work order as Assigned
                    add_dict = {'geometry': geom, 
                                'attributes': {'AssignmentStatus': 'Assigned',
                                                'AssignmentType': 'PE Inspection',
                                                'AssignmentDueDate': pe_next_insp,
                                                'AssetArea': area, 
                                                'AssetLocation': location,
                                                'AssetDescription': desc,
                                                'AssetDescriptionNotes': notes,
                                                'HazardID': hazard_id}}

                    # create new work order with same geometry as asset
                    try:      
                        print('Adding new PE work order...Hazard ID {}'.format(hazard_id))
                        work_orders.edit_features(adds = [add_dict])
                    except Exception as e:
                        print('Failed to add new work order for Hazard ID {}'.format(str(hazard_id)))
                        print(e)
                    # adds_l.append(add_dict)

                    # add attachments
                    new_work_order = work_orders.query('HazardID = {}'.format(hazard_id)).features[0]
                    new_oid = new_work_order.attributes['OBJECTID']

                    if len(assets.attachments.get_list(oid = obj_id)) > 0:
                        with tempfile.TemporaryDirectory() as dirpath:
                            try:
                                print('Downloading attachments...')
                                paths = assets.attachments.download(oid = obj_id, save_path = dirpath)
                            except Exception as e:
                                print('Failed to download attachment from Hazard ID {0}, Asset OID {1}'.format(str(hazard_id), str(obj_id)))
                                print(e)
                            for path in paths:
                                try:
                                    print('Adding attachments...')
                                    work_orders.attachments.add(oid = new_oid, file_path = path)
                                except Exception as e:
                                    print('Failed to upload attachments for Hazard ID {0}, Work Order OID {1}'.format(str(hazard_id), str(new_oid)))

            else:
                # due date set and a work order exists
                if pe_work_order_dict[hazard_id][4] == 'Completed': # and pe_work_order_dict[hazard_id][1] != None
                    print('PE work order is complete...')

                    # work order is completed
                    duedate = datetime.fromtimestamp(pe_next_insp/1000)

                    # record work order as complete in Excel
                    print('Formatting complete date...')
                    complete_date_agol = pe_work_order_dict[hazard_id][1]
                    complete_date = datetime.fromtimestamp(complete_date_agol/1000)
                    complete_date_format = complete_date.strftime("%Y-%m-%d %H:%M:%S")

                    worker = pe_work_order_dict[hazard_id][2]

                    print('Adding to completed dictionary...')
                    pe_completed[hazard_id] = [area, location, desc, notes, complete_date_format, worker, hazard_id, prod_acc]

                    # calculate next inspection date with frequency
                    print('Calculating next due date and days till due...')
                    next_due_date_agol = pe_next_insp + (pe_freq*86400000)
                    next_due_date = datetime.fromtimestamp(next_due_date_agol/1000)
                    days_till_due = (next_due_date - duedate).days

                    # if next inspection date is within a year and 10 days 
                    # create new work order with Assigned status
                    # add to upcoming in Excel
                    if days_till_due <= 375:
                        print('Creating new PE work order...')

                        # add to upcoming in Excel
                        next_due_date_format = next_due_date.strftime("%Y-%m-%d %H:%M:%S")
                        if pe_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(pe_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to upcoming dictionary...')
                        pe_upcoming[hazard_id] = [area, location, desc, notes, next_due_date_format, pe_freq, last_insp_format,
                                                    hazard_id, prod_acc]

                        # create work order as Assigned
                        add_dict = {'geometry': geom, 
                                    'attributes': {'AssignmentStatus': 'Assigned',
                                                'AssignmentType': 'PE Inspection',
                                                'AssignmentDueDate': next_due_date_agol,
                                                'AssetArea': area, 
                                                'AssetLocation': location,
                                                'AssetDescription': desc,
                                                'AssetDescriptionNotes': notes,
                                                'HazardID': hazard_id}}

                        # create new work order with same geometry as asset
                        try:      
                            print('Adding new PE work order...Hazard ID {}'.format(hazard_id))
                            work_orders.edit_features(adds = [add_dict])
                        except Exception as e:
                            print('Failed to add new work order for Hazard ID {}'.format(str(hazard_id)))
                            print(e)

                        # add attachments
                        new_work_order = work_orders.query('HazardID = {0} AND AssignmentDueDate = {1}'.format(hazard_id, next_due_date_agol)).features[0]
                        new_oid = new_work_order.attributes['OBJECTID']

                        if len(assets.attachments.get_list(oid = obj_id)) > 0:
                            with tempfile.TemporaryDirectory() as dirpath:
                                try:
                                    print('Downloading attachments...')
                                    paths = assets.attachments.download(oid = obj_id, save_path = dirpath)
                                except Exception as e:
                                    print('Failed to download attachment from Hazard ID {0}, Asset OID {1}'.format(str(hazard_id), str(obj_id)))
                                    print(e)
                                for path in paths:
                                    try:
                                        print('Adding attachments...')
                                        work_orders.attachments.add(oid = new_oid, file_path = path)
                                    except Exception as e:
                                        print('Failed to upload attachments for Hazard ID {0}, Work Order OID {1}'.format(str(hazard_id), str(new_oid)))

                        # update InspectDate and NextpeInspect with CompleteDate and new AssignmentDueDate
                        print('Updating last and next inspection dates in Assets...')
                        asset_edit = (assets.query('HazardID = {}'.format(hazard_id))).features[0]
                        asset_edit.attributes['InspectDate'] = complete_date_agol
                        asset_edit.attributes['NextPEInspect'] = next_due_date_agol
                        assets.edit_features(updates = [asset_edit])

                    else:
                        # work order is not within 40 days
                        complete_date_agol = pe_work_order_dict[hazard_id][1]
                        next_due_date_agol = pe_next_insp + (pe_freq*86400000)

                        # update InspectDate and NextPEInspect with CompleteDate and new AssignmentDueDate
                        print('Updating last and next inspection dates in Assets...')
                        asset_edit = (assets.query('HazardID = {}'.format(hazard_id))).features[0]
                        asset_edit.attributes['InspectDate'] = complete_date_agol
                        asset_edit.attributes['NextPEInspect'] = next_due_date_agol
                        assets.edit_features(updates = [asset_edit])

                else:
                    wo_obj_id = pe_work_order_dict[hazard_id][5]

                    orig_duedate = pe_work_order_dict[hazard_id][3]
                    duedate = datetime.fromtimestamp(orig_duedate/1000)
                    duedate_format = duedate.strftime("%Y-%m-%d %H:%M:%S")

                    if duedate < today:
                        print('PE work order is overdue...')

                        print('Calculating days overdue...')
                        days_overdue = (today - duedate).days

                        # record work order as overdue in Excel
                        if pe_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(pe_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to overdue dictionary...')
                        pe_overdue[hazard_id] = [duedate_format, days_overdue, area, location, desc, notes, pe_freq, last_insp_format, hazard_id, prod_acc] 
                                        
                        # update AssignmentStatus as Overdue in existing work order
                        print('Updating status to overdue...')
                        work_order_edit = (work_orders.query('OBJECTID = {}'.format(wo_obj_id))).features[0]
                        work_order_edit.attributes['AssignmentStatus'] = 'Overdue'
                        work_orders.edit_features(updates = [work_order_edit])

                    else:
                        print('PE work order is upcoming...')
                        # record work order as upcoming in Excel
                        if pe_last_insp != None:
                            last_insp_format = (datetime.fromtimestamp(pe_last_insp/1000)).strftime("%Y-%m-%d %H:%M:%S")
                        else:
                            last_insp_format = 'No previous inspection'
                        print('Adding to upcoming dictionary...')
                        pe_upcoming[hazard_id] = [area, location, desc, notes, duedate_format, pe_freq, last_insp_format, 
                                                    hazard_id, prod_acc]

    printLog('{} Operator inspections are upcoming...'.format(str(len(op_upcoming))))
    printLog('{} Operator inspections are complete...'.format(str(len(op_completed))))
    printLog('{} Operator inspections are overdue...'.format(str(len(op_overdue))))
    printLog('{} PE inspections are upcoming...'.format(str(len(pe_upcoming))))
    printLog('{} PE inspections are complete...'.format(str(len(pe_completed))))
    printLog('{} PE inspections are overdue...'.format(str(len(pe_overdue))))

    return(op_upcoming, op_completed, op_overdue, pe_upcoming, pe_completed, pe_overdue)


def moveToListService(portal, un, pw, item_id, wo_index, list_index):

    gis = GIS(portal, un, pw)
    printLog('Connected to Portal as {}'.format(gis.properties['user']['username']))

    printLog('Grabbing feature services...')
    sdi_item = gis.content.get(item_id)
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

    print('Deleting all rows from list...')
    list_lyr.edit_features(deletes = del_feat)
    print('Appending updated work orders...')
    list_lyr.edit_features(adds = copy_feat)


def cleanUp(archive_fldr, wb_path):
    '''Moves workbook to archived folder. '''

    wb_name = path.basename(wb_path)

    printLog('moving workbook to archived...')
    shutil.move(wb_path, archive_fldr + r'\{}'.format(wb_name))

    printLog('Successfully moved to "archived"!')

    archived_files = os.listdir(archive_fldr)
    if len(archived_files) > 8:

        printLog('Cleaning up archive folder...')

        oldestxlsx = min([archive_fldr + r'\{}'.format(f) for f in archived_files if f.endswith('.xlsx')], key = path.getctime)

        if oldestxlsx != archive_fldr + r'\{}'.format(wb_name):
            os.remove(oldestxlsx)


if __name__ == "__main__":

    # ------------------------------------------ maintain log file ------------------------------------------
    current = datetime.now()
    logfile = (path.abspath(path.join(path.dirname(__file__), '..', r'logs\asset-workorder-updates-ljb-LOG_{0}_{1}.txt'.format(current.month, current.year))))
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
    asset_index = json_ago["assetindex"]
    wo_index = json_ago["workorderindex"]
    list_index = json_ago["listindex"]

    # ------------------------------------------ get email creds ------------------------------------------
    email_text = open(workingfldr + r'\email-info.json').read()
    json_email = json.loads(email_text)

    email_recipient = json_email['email_to']
    email_from = json_email['email_from']
    email_cc = json_email['email_cc']
    email_subject = json_email['subject']
    email_message = json_email['message']

    # ------------------------------------------ excel workbook columns ------------------------------------
    op_complete_cols = [['A1','Area'],['B1', 'Location'],['C1', 'Description'],['D1', 'Asset Description Notes'],
                        ['E1', 'Completion Date'],['F1', 'Who Completed Inspection'],['G1', 'Asset ID of Hazard'],
                        ['H1', 'Production Accessible'],['I1', 'Type (Crane or WWS)']]

    pe_complete_cols = [['A1','Area'],['B1', 'Location'],['C1', 'Description'],['D1', 'Asset Description Notes'],
                        ['E1', 'Completion Date'],['F1', 'Who Completed Inspection'],['G1', 'Asset ID of Hazard'],
                        ['H1', 'Production Accessible']]

    op_upcoming_cols = [['A1', 'Area'],['B1', 'Location '],['C1', 'Description'],['D1', 'Asset Description Notes'],
                        ['E1', 'Due Date'],['F1', 'Frequency of Inspection'],['G1', 'Last Inspection Date'],
                        ['H1', 'Asset ID (Hazard  ID)'],['I1', 'Production Accessible'],['J1', 'Type (Crane or WWS)']]

    pe_upcoming_cols = [['A1', 'Area'],['B1', 'Location '],['C1', 'Description'],['D1', 'Asset Description Notes'],
                        ['E1', 'Due Date'],['F1', 'Frequency of Inspection'],['G1', 'Last Inspection Date'],
                        ['H1', 'Asset ID (Hazard  ID)'],['I1', 'Production Accessible']]

    op_overdue_cols = [['A1','Original Due Date'],['B1', 'Days Overdue'],['C1', 'Area'],['D1', 'Location'],
                        ['E1', 'Description'],['F1', 'Asset Description Notes'],['G1', 'Frequency'],
                        ['H1', 'Last Inspected Date'],['I1', 'Asset ID of Hazard'],['J1', 'Production Accessible'],
                        ['K1', 'Type (Crane or WWS)']]

    pe_overdue_cols = [['A1','Original Due Date'],['B1', 'Days Overdue'],['C1', 'Area'],['D1', 'Location'],
                        ['E1', 'Description'],['F1', 'Asset Description Notes'],['G1', 'Frequency'],
                        ['H1', 'Last Inspected Date'],['I1', 'Asset ID of Hazard'],['J1', 'Production Accessible']]  
    
    # ------------------------------------------ execute functions ------------------------------------------
    try: 
        printLog('Grabbing services and sorting data...')
        asset, work_orders, asset_dict, \
            op_work_order_dict, pe_work_order_dict = buildQueryDictionaries(orgURL, username, password, itemID, asset_index, wo_index)
        printLog('Updating assets and work orders...')
        op_upcom, op_com, op_overdue, \
            pe_upcom, pe_com, pe_overdue = updateServices(asset, work_orders, asset_dict, op_work_order_dict, pe_work_order_dict, current)
        printLog('Moving work orders to list...')
        moveToListService(orgURL, username, password, itemID, wo_index, list_index)

        worksheet_dict = {'Operator Completed': [op_com, op_complete_cols], 
                        'PE Completed': [pe_com, pe_complete_cols],
                        'Operator Upcoming': [op_upcom, op_upcoming_cols],
                        'PE Upcoming': [pe_upcom, pe_upcoming_cols],
                        'Operator Overdue': [op_overdue, op_overdue_cols], 
                        'PE Overdue': [pe_overdue, pe_overdue_cols]}
        printLog('Creating Excel workbook...')
        wb, wb_path = createWorkbook(workingfldr)
        printLog('Populating worksheets...')
        addWorksheet(wb, worksheet_dict)
        wb.close()

        printLog('Sendning e-mail...')
        sendemail(email_recipient, email_from, email_subject, email_message, email_cc, att=wb_path)

        printLog('Clean up, aisle 9...')
        cleanUp(archfldr, wb_path)

        printLog("Success! \n ------------------------------------ \n\n")

    except Exception: 
        logging.error("EXCEPTION OCCURRED", exc_info=True)
        logging.info('Sending email to LJB staff for inspection... ')
        email_err_subject = "SDI Caster Script Error"
        email_err_message = "An error occurred while updating the WWS Assignments... \n" \
                            + "See the attached log file for error message.\n" 
        sendemail(eto=email_cc, efrom=email_from, subject=email_err_subject, message=email_err_message, att=logfile)
        printLog("Quitting! \n ------------------------------------ \n\n")  


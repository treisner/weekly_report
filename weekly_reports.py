import win32com.client
import os
import pymssql
import re
import pandas as pd
import django
from django.conf import settings
from django.core.mail import EmailMessage

settings.configure(DEBUG=True,
                   EMAIL_USE_TLS=True,
                   EMAIL_BACKEND ='django.core.mail.backends.smtp.EmailBackend',
                   EMAIL_HOST = 'smtp-relay.sendinblue.com',
                   EMAIL_PORT = 587,
                   EMAIL_HOST_USER = 'george@treisner.com',
                   EMAIL_HOST_PASSWORD = 'N7IypPgsaLZOztcG',
                   )
django.setup()
from_email = 'reports_noreply@sos-analytics.org'

testing = False  # Reports for district "NY Metro" only
quicktest = False  # District level only, no territory sheets.  Only used if testing is True
really_quick_test = False #Don't generate any reports
send_email = True #if False then files generated, but no email sent

region_reports = True
district_reports = True
territory_reports = True

missing_email = None
testing_email = ['george@treisner.com']
all_to_testing_email = False  # Override all emails and use testing_email instead.  This only is used when testing if False
email_header = ''
if all_to_testing_email:
    email_header = "<h3>This email is for testing and is only being sent to these recipients:<ul>"
    for addr in testing_email:
        email_header = email_header + '<li>' + addr \
                       + '</li>'
    email_header = email_header + '</ul></h3>'


region_cc = ['george@treisner.com', 'pbruzik@sos-analytics.org',
             'TCentofanti@mannkindcorp.com', 'vkwok@mannkindcorp.com',
             'GLeviness@mannkindcorp.com', 'mBazzani@mannkindcorp.com',
             'AGalindo@mannkindcorp.com', 'Rshea@mannkindcorp.com', 'agoyal@mannkindcorp.com',
             'zxi@mannkindcorp.com', 'ffungo@mannkindcorp.com',
             ]



if testing: print("*****TESTING*****")
if all_to_testing_email:
    print("*****All emails go to: ", testing_email)
if region_reports: print("Region reports will be generated")
if district_reports: print("District reports will be generated")
if territory_reports: print("Territory reports will be generated")




# # Get current roster file
# OneDrive = r'C:\Users\treisner\OneDrive - Treetops Analytics LLC\Email attachments from Flow'
# dir = os.listdir(OneDrive)
# roster_file_names = [d for d in dir if re.match(r'Sales.Roster.Alignment.(\d+)_(\d+)_(\d\d\d\d)\.xlsx',d)]
# # roster_file_names = [d for d in dir if re.match(r'Mannkind Roster-*',d)]
#
# roster_files = [(os.path.getctime('{OneDrive}\{file_name}'.format(OneDrive=OneDrive,file_name=file_name)),'{OneDrive}\{file_name}'.format(OneDrive=OneDrive,file_name=file_name)) for file_name in roster_file_names]
# current_roster=max(roster_files)[1]
# # roster=pd.read_excel(current_roster,sheet_name='Field-Roster-7123')
# roster=pd.read_excel(current_roster,sheet_name='Roster')
# print('Using Roster: ',current_roster)

conn = pymssql.connect(server='localhost', database='symphony')
cursor = conn.cursor()
xl = win32com.client.DispatchEx('Excel.Application')

report_dir = r"c:\users\treisner\documents\mannkind\reports"
template_dir = r"c:\users\treisner\documents\mannkind\weekly report"
template_name = r"rank report template.xlsx"
template_book = xl.Workbooks.Open(Filename=template_dir + '\\' + template_name)
template_sheet = template_book.Sheets("Template")
template_sheet_NTW = template_book.Sheets("NTW_Template")
#template_sheet_OPTW = template_book.Sheets("OPTW_Template")
district_selection = r'[afrezza trx]>0'
district_geography = r"district='{dist}'"
territory_selection = r'[Terr Rank]<10000'
territory_geography = r"Terr='{terr}'"
region_geography = r"region='{region}'"
region_selection = r'[afrezza trx]>0'

with open(template_dir + '\\region template.html') as f:
    region_text = f.read()

with open(template_dir + '\\district template.html') as f:
    district_text = f.read()

with open(template_dir + '\\territory template.html') as f:
    territory_text = f.read()

if all_to_testing_email:
    region_text = email_header + region_text
    district_text = email_header + district_text
    territory_text = email_header + territory_text

sql = "select * from report_months order by month"
cursor.execute(sql)
month = []
months = ''
month_list = cursor.fetchall()
for m in month_list:
    month.append(m[0])
    months += '[Afrezza {}], '.format(m[0])

for c in template_sheet.UsedRange:
    afrezza_header = re.match(r'[a|A]frezza [c|C](\d+)', c.Value)
    if afrezza_header:
        index = int(afrezza_header.group(1))
        c.Value = "Afrezza {}".format(month[12 - index])

sql_report_template = """
SELECT Region, District, Terr as Territory,  [Region Rank], [District Rank],  [Terr Rank],[View_calls_in_date_range].calls, pdrp_indicator, 
LastName,FirstName, MiddleInitial,  TitleOrSuffix, Address1, Address2, City, StateAbbr, ''''+ZipCode, Specialty1, Tier,  Share, trend, trend_rapid, Afrezza_trx_this_week, 
[Afrezza TRX],  [RA TRX],  {months} ''''+npi, ''''+data_week 
FROM ranked_detail LEFT JOIN View_calls_in_date_range ON ranked_detail.REL_ID = View_calls_in_date_range.PrescriberPracRelID
where {geography} and pdrp_indicator is null and {selection}
order by [region Rank];
"""

sql_report_template_no_pdrp = """
SELECT Region, district, Terr, Null AS [Region Rank], Null AS [District Rank], Null AS [Terr Rank],[View_calls_in_date_range].calls, pdrp_indicator, 
LastName,FirstName, MiddleInitial,  TitleOrSuffix, Address1, Address2, City, StateAbbr, ''''+ZipCode, Specialty1, Tier,  Share, trend, trend_rapid, Null as [Afrezza_trx_this_week], 
Null AS [Afrezza TRX], Null AS [RA TRX], NULL AS [Afrezza 201610], NULL AS [Afrezza 201611], NULL AS [Afrezza 201612], NULL AS [Afrezza 201701], NULL AS [Afrezza 201702], NULL AS [Afrezza 201703], NULL AS [Afrezza 201704], NULL AS [Afrezza 201705], NULL AS [Afrezza 201706], NULL AS [Afrezza 201707], NULL AS [Afrezza 201708], NULL AS [Afrezza 201709], ''''+npi, ''''+data_week 
FROM ranked_detail LEFT JOIN View_calls_in_date_range ON ranked_detail.REL_ID = View_calls_in_date_range.PrescriberPracRelID
where {geography} and pdrp_indicator='Y' and {selection};
"""

sql_report_NTW = "Select * from new_this_week where {geography} "
sql_report_OPTW = "Select * from old_product_this_week where {geography} "

def make_workbook(workbook_name, worksheet_name, copy_sheet):
    file_name = report_dir + '\\' + workbook_name + r'.xlsx'
    try:
        os.remove(file_name)
    except:
        pass
    wb = xl.Workbooks.Add()
    copy_sheet.Copy(None, wb.Sheets(wb.Sheets.Count))
    new_sheet = wb.Sheets(wb.Sheets.Count)
    new_sheet.Name = worksheet_name
    ws = wb.Sheets('Sheet1')
    ws.Delete()
    wb.SaveAs(file_name)
    return wb


def report_to_worksheet(workbook_name, worksheet_name, geography, selection):
    sql = sql_report_template.format(geography=geography, selection=selection, months=months)
    # print(sql)
    # quit()
    cursor.execute(sql)
    results = cursor.fetchall()
    i=0
    ws = workbook_name.Sheets(worksheet_name)
    if really_quick_test:
        return ws
    for i, row in enumerate(results):
        ws.Range(ws.Cells(i + 2, 1), ws.Cells(i + 2, len(row))).Value = row
    offset = i + 3
    sql = sql_report_template_no_pdrp.format(geography=geography, selection=selection)
    cursor.execute(sql)
    results = cursor.fetchall()
    ws = workbook_name.Sheets(worksheet_name)
    for i, row in enumerate(results):
        ws.Range(ws.Cells(i + offset, 1), ws.Cells(i + offset, len(row))).Value = row
        workbook_name.Save()
    return ws


def report_NTW_to_worksheet(worksheet_name, geography, workbook):
    if really_quick_test:
        return
    sql = sql_report_NTW.format(geography=geography)
    cursor.execute(sql)
    results = cursor.fetchall()
    if len(results) > 0:
        template_sheet_NTW.Copy(None, workbook.Sheets(workbook.Sheets.Count))
        new_sheet = workbook.Sheets(workbook.Sheets.Count)
        new_sheet.Name = worksheet_name
        for i, row in enumerate(results):
            new_sheet.Range(new_sheet.Cells(i + 2, 1), new_sheet.Cells(i + 2, len(row))).Value = row
        workbook.Save()
    return


# def report_OPTW_to_worksheet(worksheet_name, geography, workbook):
#     if really_quick_test:
#         return
#     sql = sql_report_OPTW.format(geography=geography)
#     cursor.execute(sql)
#     results = cursor.fetchall()
#     if len(results) > 0:
#         template_sheet_OPTW.Copy(None, workbook.Sheets(workbook.Sheets.Count))
#         new_sheet = workbook.Sheets(workbook.Sheets.Count)
#         new_sheet.Name = worksheet_name
#         for i, row in enumerate(results):
#             new_sheet.Range(new_sheet.Cells(i + 2, 1), new_sheet.Cells(i + 2, len(row))).Value = row
#         workbook.Save()
#     return


sql = "select week_ending from view_data_week where week_number=52"
cursor.execute(sql)
dw = cursor.fetchall()[0][0]

data_week = '{}-{}-{}'.format(dw[0:4], dw[4:6], dw[6:8])

print(data_week)

if testing:
    sql = "select region, [reg number] from zipterr where region='East' group by region,[reg number] order by region"
else:
    sql = "select region, [reg number] from zipterr where region<>'*Unassigned' group by region,[reg number] order by region"
#    sql = "select region, [reg number] from zipterr where region='West' group by region,[reg number] order by region"
cursor.execute(sql)
regions = cursor.fetchall()
if not regions:
    regions=["National","10000"]
for k, reg in enumerate(regions):
    #if k==0: continue
    region = reg[0]
    print(k, reg)
    region_file_name = 'MannKind region {} {}'.format(data_week, region)
    if testing:
        sql = "select district, [dist number] from zipterr where district<>'*Unassigned' and district='NY Metro' group by district, [dist number] order by district"
    else:
        sql = "select district, [dist number]  from zipterr where district<>'*Unassigned' and region='{}'  group by district, [dist number] order by district".format(
            region)

    # sql = "select district, [dist number]  from zipterr where district<>'*Unassigned' and region='{}'  group by district, [dist number] order by district".format(region)

    cursor.execute(sql)
    districts = cursor.fetchall()
    rwb = make_workbook(region_file_name, region, template_sheet)
    rws = report_to_worksheet(rwb, region, region_geography.format(region=region), region_selection)
    report_NTW_to_worksheet("New This Week", region_geography.format(region=region), rwb)
    #report_OPTW_to_worksheet("Sunsetted This Week", region_geography.format(region=region), rwb)
    for j, district in enumerate(districts):
        print("     District", district)
        dist = 'District {}'.format(district[0].replace(r'/', '-'))
        district_file_name = 'MannKind District {} {}'.format(data_week, district[0].replace(r'/', '-'))

        sql = "select terr, [terr number] from zipterr where district='{}' group by terr, [terr number] order by terr".format(
            district[0])
        cursor.execute(sql)
        terrs = cursor.fetchall()
        print(terrs)
        dwb = make_workbook(district_file_name, dist, template_sheet)
        dws = report_to_worksheet(dwb, dist, district_geography.format(dist=district[0]), district_selection)
        report_NTW_to_worksheet("New This Week", district_geography.format(dist=district[0]), dwb)
        #report_OPTW_to_worksheet("Sunsetted This Week", district_geography.format(dist=district[0]), dwb)
        dws.Copy(None, rwb.Sheets(rwb.Sheets.Count))
        rdws = rwb.Sheets(rwb.Sheets.Count)
        rdws.Name = dist
        for i, terr in enumerate(terrs):
            if testing and quicktest: break
            territory = terr[0].replace(r'/', '-')
            territory_file_name = "MannKind Territory {} {}-{}".format(data_week, district[0].replace(r'/', '-'),
                                                                       territory)
            print("          ", i, terr)
            template_sheet.Copy(None, dwb.Sheets(dwb.Sheets.Count))
            tws = dwb.Sheets(dwb.Sheets.Count)
            tws.Name = territory
            ws = report_to_worksheet(dwb, territory, territory_geography.format(terr=terr[0]), territory_selection)
            #print(territory, territory_geography.format(terr=terr[0]), territory_selection)
            if territory_reports:
                twb = make_workbook(territory_file_name, territory, ws)
                report_NTW_to_worksheet("New This Week", territory_geography.format(terr=terr[0]), twb)
                #report_OPTW_to_worksheet("Sunsetted This Week", territory_geography.format(terr=terr[0]), twb)
                twb.Sheets(1).Activate()
                twb.Save()
                twb.Close()
                sql = "select EmailAddress from TP_RosterAlignment where SalesAreaLevel='Terr' and SalesAreaLevel1Code ='{}'".format(
                    terr[1])
                cursor.execute(sql)
                try:
                    terr_email = [cursor.fetchall()[0][0]]
                except:
                    terr_email = None
                # try:
                #     terr_email_roster = roster[roster['Area ID']==int(terr[1])]['Business Email'].values[0]
                # except:
                #     terr_email_roster = missing_email
                # terr_email = terr_email_roster # roster overides veeva
                if testing or all_to_testing_email: terr_email = testing_email
                if terr_email and not (all_to_testing_email and i + j + k > 0):

                    subject = "Territory {} Writers for Week Ending {}".format(terr[0], data_week)
                    contents = territory_text.replace(r'{territory}', terr[0]).replace(r'{data_week}',
                                                                                       data_week).replace('\n', "")
                    print(terr_email)
                    if send_email:
                        print("          TERRITORY: Sending to", terr_email)
                        email = EmailMessage(
                            subject=subject,
                            body=contents,
                            from_email=from_email,
                            to=terr_email,
                        )
                        email.attach_file(report_dir + '\\' + territory_file_name + r'.xlsx')
                        email.send()

        dwb.Sheets(1).Activate()
        dwb.Save()
        dwb.Close()
        sql = "select EmailAddress from TP_RosterAlignment where SalesAreaLevel='Dist' and SalesAreaLevel2Code ='{}'".format(
            district[1])
        cursor.execute(sql)
        try:
            district_email = [cursor.fetchall()[0][0]]
        except:
            district_email = None

        if testing or all_to_testing_email: district_email = testing_email
        if district_email and district_reports and not (all_to_testing_email and j + k > 0):

            subject = "District {} Writers for Week Ending {}".format(district[0], data_week)
            contents = district_text.replace(r'{district}', district[0]).replace(r'{data_week}', data_week).replace(
                '\n', "")
            print(district_email)
            if send_email:
                print("     DISTRICT: Sending to", district_email)
                email = EmailMessage(
                    subject=subject,
                    body=contents,
                    from_email=from_email,
                    to=district_email,
                )
                email.attach_file(report_dir + '\\' + district_file_name + r'.xlsx')
                email.send()

    rwb.Sheets(1).Activate()
    rwb.Save()
    rwb.Close()
    # sql = "select EmailAddress from TP_RosterAlignment where SalesAreaLevel='REGN' and SalesAreaLevel3Code ='{}'".format(
    #     reg[1])
    # cursor.execute(sql)
    # try:
    #     region_email = [cursor.fetchall()[0][0]]
    # except:
    #     region_email = ['jmainland@mannkindcorp.com']
    region_email = ['jmainland@mannkindcorp.com']
    # region_email = ['george@treisner.com']

    print(region_email)
    if testing or all_to_testing_email:
        region_email = testing_email
        region_cc = []
    if region_email and region_reports:
        print("REGION: Sending to", region_email)
        print("REGION CC:", region_cc)
        subject = "Region {} Writers for Week Ending {}".format(region, data_week)
        contents = region_text.replace(r'{region}', region).replace(r'{data_week}', data_week).replace('\n', "")
        print(region_email)
        if send_email:
            print("     Region: Sending to", region_email)
            email = EmailMessage(
                subject=subject,
                body=contents,
                from_email=from_email,
                to=region_email,
                cc=region_cc
            )
            email.attach_file(report_dir + '\\' + region_file_name + r'.xlsx')
            email.send()

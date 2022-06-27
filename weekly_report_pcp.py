import os
import re
import yagmail
import pandas as pd
from win32com import client
from sqlalchemy import create_engine, MetaData, Table, func, desc
from sqlalchemy.orm import mapper, sessionmaker, relationship
from sqlalchemy import Column, Float, Unicode, DateTime
from sqlalchemy.ext.declarative import declarative_base


Base = declarative_base()
metadata = Base.metadata
run_national_reports = True
run_district_reports = False
run_territory_reports = False
send_email = True #if False then files generated, but no email sent
testing = False #emails, if any, all go to testing_address
testing_address = 'george@treisner.com'

national_email_list = ['pbruzik@sos-analytics.org','george@treisner.com', 'jmainland@mannkindcorp.com']

xl = client.DispatchEx('Excel.Application')
template_book = xl.Workbooks.Open(
    r'C:\Users\treisner\Documents\MannKind\Weekly Report\weekly_report_pcp_template.xlsx')
template_sheet = template_book.Sheets("Template")
yag = yagmail.SMTP("mannkind.reports@gmail.com", 'sos.tta01')
yag.useralias = "MannKind Reports"

def main():
    session = loadSession()
    roster = Roster()
    print('Using Roster: ', roster.current_roster)

    dw = session.query(DataWeek.week_ending).all()[0][0]
    data_week = f'{dw[0:4]}-{dw[4:6]}-{dw[6:8]}'

    pcps = session.query(
        PcpCallPlan.district.label('district'),
        PcpCallPlan.terr.label('territory'),
        PcpCallPlan.spec_code.label('specialty'),
        PcpCallPlan.first_name.label('first_name'),
        PcpCallPlan.middle_name.label('middle_name'),
        PcpCallPlan.last_name.label('last_name'),
        PcpCallPlan.address.label('address'),
        PcpCallPlan.city.label('city'),
        PcpCallPlan.state.label('state'),
        PcpCallPlan.zip.label('zip_code'),
        PcpCallPlan.tier.label('tier'),
        PcpCallPlan.decile.label('long_decile'),
        AfrezzaLast13Weeks.trx.label('trx'),
        AfrezzaLast13Weeks.last_week.label('last_rx_week'),
        PcpCallPlan.pdrp_indicator.label('pdrp')
    ).outerjoin(AfrezzaLast13Weeks,
                AfrezzaLast13Weeks.rel_id == PcpCallPlan.rel_id) \
        .order_by(desc('trx'), desc('long_decile'), 'territory', 'last_name').all()

    districts = session.query(Zipterr.district, Zipterr.dist_number).group_by(Zipterr.district,
                                                                              Zipterr.dist_number).all()
    territories = session.query(Zipterr.territory, Zipterr.terr_number).group_by(Zipterr.territory,
                                                                                 Zipterr.terr_number).all()
    for d in districts:
        if d[0] and run_district_reports:
            file_name = f'C:\\Users\\treisner\\Documents\\MannKind\\reports\\district-{d[0]}-{data_week}.xlsx'
            report_for_area(d[0], d[1], roster, file_name, pcps, level='district')

    for t in territories:
        if t[0] and run_territory_reports:
            file_name = f'C:\\Users\\treisner\\Documents\\MannKind\\reports\\territory-{t[0]}-{data_week}.xlsx'
            report_for_area(t[0], t[1], roster, file_name, pcps, level='territory')

    if run_national_reports:
        file_name = f'C:\\Users\\treisner\\Documents\\MannKind\\reports\\national-{data_week}.xlsx'
        report_for_area(None, None, roster, file_name, pcps, level='national')

    template_book.Close()
    yag.close()


def report_for_area(area_name, area_id, roster, file_name, pcps, level='national'):
    email_address = roster.get_email(area_id)

    if not area_name:
        area_name = 'national'
        email_address = national_email_list
    print(f'{level}={area_name}\t{email_address=}')
    wb = make_workbook(xl, file_name, area_name, template_sheet)
    if level == 'territory':
        report_to_worksheet(wb, area_name, pcps, district=None, territory=area_name)
    if level == 'district':
        report_to_worksheet(wb, area_name, pcps, district=area_name, territory=None)
    if level == 'national':
        report_to_worksheet(wb, 'national', pcps, district=None, territory=None)
    if testing:
        email_address = testing_address
        print(f'TESTING: emails go to {email_address}')
    if send_email and email_address:
        yag.send(email_address, 'PCP Rx Report',
                 'The attached shows Rx activity for the PCP sales force.',
                 attachments=file_name)


def report_to_worksheet(workbook_name, worksheet_name, report_data, district=None, territory=None):
    offset = 2
    ws = workbook_name.Sheets(worksheet_name)
    for row in report_data:
        if district and row[0] == district:
            ws.Range(ws.Cells(offset, 1), ws.Cells(offset, len(row))).Value = row
            offset += 1

        if territory and row[1] == territory:
            ws.Range(ws.Cells(offset, 1), ws.Cells(offset, len(row))).Value = row
            offset += 1

        if not territory and not district:
            ws.Range(ws.Cells(offset, 1), ws.Cells(offset, len(row))).Value = row
            offset += 1

    workbook_name.Save()
    workbook_name.Close()
    return


def make_workbook(xl, file_name, worksheet_name, copy_sheet):
    try:
        os.remove(file_name)
    except FileNotFoundError:
        pass
    wb = xl.Workbooks.Add()
    copy_sheet.Copy(None, wb.Sheets(wb.Sheets.Count))
    new_sheet = wb.Sheets(wb.Sheets.Count)
    new_sheet.Name = worksheet_name
    ws = wb.Sheets('Sheet1')
    ws.Delete()
    wb.SaveAs(file_name)
    return wb


class Calls(Base):
    __tablename__ = 'Mannkind_calls'
    call_id = Column(Unicode(255), primary_key=True, unique=True)
    call_name = Column(Unicode(255))
    account_id = Column(Unicode(255))
    call_status = Column(Unicode(255))
    call_activity_type = Column(Unicode(255))
    call_date_time = Column(DateTime)
    call_type = Column(Unicode(255))
    territory = Column(Unicode(255))
    rep_id = Column(Unicode(255))
    rep_first_name = Column(Unicode(255))
    rep_last_name = Column(Unicode(255))
    attendee_first_name = Column(Unicode(255))
    attendee_last_name = Column(Unicode(255))
    address_line1 = Column(Unicode(255))
    address_line2 = Column(Unicode(255))
    city = Column(Unicode(255))
    state = Column(Unicode(255))
    zip_code = Column(Unicode(255))
    location_id = Column(Unicode(255))
    credentials = Column(Unicode(255))
    specialty = Column(Unicode(255))
    npi = Column(Unicode(255))
    me = Column(Unicode(255))
    detail_1 = Column(Unicode(255))
    detail_2 = Column(Unicode(255))
    detail_3 = Column(Unicode(255))
    detail_4 = Column(Unicode(255))
    detail_5 = Column(Unicode(255))
    detail_6 = Column(Unicode(255))
    detail_7 = Column(Unicode(255))
    detail_8 = Column(Unicode(255))
    detail_9 = Column(Unicode(255))
    detail_10 = Column(Unicode(255))
    parent_call_id = Column(Unicode(255))
    custom_1 = Column(Unicode(255))
    custom_2 = Column(Unicode(255))
    custom_3 = Column(Unicode(255))
    custom_4 = Column(Unicode(255))
    custom_5 = Column(Unicode(255))

class DataWeek(Base):
    __tablename__ = 'view_data_week'
    week_number = Column(Float, primary_key=True, unique=True)
    week_ending = Column(Unicode(8))


class Zipterr(Base):
    __tablename__ = 'zipterr_pcp'

    zip = Column(Unicode(255), primary_key=True, unique=True)
    territory = Column(Unicode(255))
    terr_number = Column('terr number', Unicode(255), index=True)
    district = Column(Unicode(255))
    dist_number = Column('dist number', Unicode(255), index=True)


class AfrezzaLast13Weeks(Base):
    __tablename__ = 'Afrezza_by_Rel_id_last_13_weeks'

    rel_id = Column(Unicode(255), primary_key=True, unique=True)
    trx = Column(Float(53))
    nrx = Column(Float(53))
    first_week = Column(Unicode(255))
    last_week = Column(Unicode(255))
    zip_code = Column(Unicode(5))
    territory = Column(Unicode(255))
    District = Column(Unicode(255))
    trx_this_week = Column(Float(53))


class PcpCallPlan(Base):
    __tablename__ = 'pcp_call_plan'

    terr_number = Column(Unicode(255))
    terr = Column(Unicode(255))
    dist_number = Column(Unicode(255))
    district = Column(Unicode(255))
    first_name = Column(Unicode(255))
    middle_name = Column(Unicode(255))
    last_name = Column(Unicode(255))
    title = Column(Unicode(255))
    spec_code = Column(Unicode(255), index=True)
    spec_description = Column(Unicode(255))
    address = Column(Unicode(255))
    city = Column(Unicode(255))
    state = Column(Unicode(255))
    zip = Column(Unicode(255))
    rel_id = Column(Unicode(255), primary_key=True, unique=True)
    dea_number = Column(Unicode(255))
    npi = Column(Unicode(255))
    ama_number = Column(Unicode(255))
    ama_check_digit = Column(Unicode(255))
    pdrp_indicator = Column(Unicode(255))
    decile = Column(Float)
    mkt_trx = Column(Float)
    tier = Column(Float)
    location_phone = Column(Unicode(255))
    location_fax = Column(Unicode(255))
    crm_id = Column(Unicode(255), index=True)
    long_decile = Column(Float)
    rapid_decile = Column(Float)
    oral_decile = Column(Float)
    cgm_decile = Column(Float)


def loadSession():
    """"""
    dbPath = 'mssql://@localhost/symphony?trusted_connection=yes&driver=ODBC+Driver+13+for+SQL+Server'
    engine = create_engine(dbPath)
    Session = sessionmaker(bind=engine)
    session = Session()
    return session


class Roster:
    """Gets the latest roster file from the roster_dir"""
    ROSTER_DIR = r'C:\Users\treisner\Documents\MannKind\Weekly Report'
    FILE_NAME_PATTERN = r'MannKind.Roster.(.*)\.xlsx'
    HEADER_ROW = 11  # zero based index of the row containing the column headers

    def __init__(self):
        roster_dir = self.ROSTER_DIR
        file_list = os.listdir(roster_dir)
        roster_file_names = [d for d in file_list if re.match(self.FILE_NAME_PATTERN, d)]
        roster_files = [(os.path.getctime(f'{roster_dir}\{file_name}'), f'{roster_dir}\{file_name}') for file_name in
                        roster_file_names]
        self.current_roster = max(roster_files)[1]
        self.roster = pd.read_excel(self.current_roster, header=self.HEADER_ROW)

    def get_email(self, terr_number):
        email = None
        if terr_number:
            try:
                email = self.roster[self.roster['Geo ID'] == terr_number]['Business Email Address'].values[0]
                if len(email) < 4:
                    email = None
            except:
                email = None
        return email


if __name__ == '__main__':
    main()

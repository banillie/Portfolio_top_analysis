'''

Programme for creating an aggregate project dashboard

input documents:
1) Dashboard master document - this is an excel file. It should have the dashboard design, with all projects structured
in the correct way (order etc), but all data fields left blank. If project data is not being put into the correct
part of the master, ensure that the project name is consistent in master data and this document. The names need to be
exactly the same for information to be released.
2) Master data for two quarters - this will usually be latest and previous quarter

output document:
3) Dashboard with all project data placed into dashboard and formatted correctly.

Instructions:
1) provide path to dashboard master
2) provide path to master data sets
3) provide path and specify file name for output document

Supplementary instructions:
These things need to be done to check and assure the data going into the dashboard. Use the seperate programme available
for undertaking these tasks.
1) Check that project stage/last at BICC data is correct
2) Check the last at / next at BICC project data is correct.
3) Make sure IPA DCA ratings are correct.


'''

from openpyxl import load_workbook
from bcompiler.utils import project_data_from_master
import datetime
from openpyxl.styles import PatternFill, Font
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import Rule, IconSet, FormatObject

'''function for putting keys of interest plus values into a list'''
def key_of_interest(project_name, data, key):
    data = data[project_name]
    cell_keys = key
    output_list = []
    for item in data.items():
        if item[0] in cell_keys:
            output_list.append(item)
    return output_list


'''function for converting dates into concatenated written time periods'''
def concatenate_dates(date):
    today = datetime.datetime(2019, 2,
                              4)  # this needs to be the date the report is being discussed at BICC. Python date format (YYYY,MM,DD)
    if date != None:
        a = (date - today.date()).days
        year = 365
        month = 30
        fortnight = 14
        week = 7
        if a >= 365:
            yrs = int(a / year)
            holding_days_years = a % year
            months = int(holding_days_years / month)
            holding_days_months = a % month
            fortnights = int(holding_days_months / fortnight)
            weeks = int(holding_days_months / week)
        elif 0 <= a <= 365:
            yrs = 0
            months = int(a / month)
            holding_days_months = a % month
            fortnights = int(holding_days_months / fortnight)
            weeks = int(holding_days_months / week)
            # if 0 <= a <=60:
        elif a <= -365:
            yrs = int(a / year)
            holding_days = a % -year
            months = int(holding_days / month)
            holding_days_months = a % -month
            fortnights = int(holding_days_months / fortnight)
            weeks = int(holding_days_months / week)
        elif -365 <= a <= 0:
            yrs = 0
            months = int(a / month)
            holding_days_months = a % -month
            fortnights = int(holding_days_months / fortnight)
            weeks = int(holding_days_months / week)
            # if -60 <= a <= 0:
        else:
            print('something is wrong and needs checking')

        if yrs == 1:
            if months == 1:
                return ('{} yr, {} mth'.format(yrs, months))
            if months > 1:
                return ('{} yr, {} mths'.format(yrs, months))
            else:
                return ('{} yr'.format(yrs))
        elif yrs > 1:
            if months == 1:
                return ('{} yrs, {} mth'.format(yrs, months))
            if months > 1:
                return ('{} yrs, {} mths'.format(yrs, months))
            else:
                return ('{} yrs'.format(yrs))
        elif yrs == 0:
            if a == 0:
                return ('Today')
            elif 1 <= a <= 6:
                return ('This week')
            elif 7 <= a <= 13:
                return ('Next week')
            elif -7 <= a <= -1:
                return ('Last week')
            elif -14 <= a <= -8:
                return ('-2 weeks')
            elif 14 <= a <= 20:
                return ('2 weeks')
            elif 20 <= a <= 60:
                if today.month == date.month:
                    return ('Later this mth')
                elif (date.month - today.month) == 1:
                    return ('Next mth')
                else:
                    return ('2 mths')
            elif -60 <= a <= -15:
                if today.month == date.month:
                    return ('Earlier this mth')
                elif (date.month - today.month) == -1:
                    return ('Last mth')
                else:
                    return ('-2 mths')
            elif months == 12:
                return ('1 yr')
            else:
                return ('{} mths'.format(months))

        elif yrs == -1:
            if months == -1:
                return ('{} yr, {} mth'.format(yrs, -(months)))
            if months < -1:
                return ('{} yr, {} mths'.format(yrs, -(months)))
            else:
                return ('{} yr'.format(yrs))
        elif yrs < -1:
            if months == -1:
                return ('{} yrs, {} mth'.format(yrs, -(months)))
            if months < -1:
                return ('{} yrs, {} mths'.format(yrs, -(months)))
            else:
                return ('{} yrs'.format(yrs))
    else:
        return ('None')


'''function for calculating if confidence has increased decreased'''
def up_or_down(name, dict_1, dict_2):
    conf_1 = dict_1[name][0][1]
    # print(conf_1)
    try:
        conf_2 = dict_2[name][0][1]
        # print(conf_2)
        if conf_1 == conf_2:
            return ('Change', int(0))
        elif conf_1 != conf_2:
            if conf_2 == 'Green':
                if conf_1 != 'Amber/Green':
                    return ('Change', int(-1))
            elif conf_2 == 'Amber/Green':
                if conf_1 == 'Green':
                    return ('Change', int(1))
                else:
                    return ('Change', int(-1))
            elif conf_2 == 'Amber':
                if conf_1 == 'Green':
                    return ('Change', int(1))
                elif conf_1 == 'Amber/Green':
                    return ('Change', int(1))
                else:
                    return ('Change', int(-1))
            elif conf_2 == 'Amber/Red':
                if conf_1 == 'Red':
                    return ('Change', int(-1))
                else:
                    return ('Change', int(1))
            else:
                return ('Change', int(1))
    except KeyError:
        return ('Change', 'NEW')


'''function for list/dictionary for current quarter'''
def making_dict(names):
    d = {}
    for x in names:
        be = key_of_interest(x, data_current_quarter, to_capture_current_quarter)
        d[x] = be
    return d


'''function for adding concatenated word strings to dictionary.
note probably don't need the above function now, but can tidy up later'''
def adding_con(d, d2):
    new_dic = {}
    for x in d:
        #print(x)
        be = key_of_interest(x, data_current_quarter, to_capture_current_quarter)
        lis = []
        for i in range(1, 5):
            try:
                y = concatenate_dates(d[x][i][1])
            except TypeError:
                y = 'None'
            b = (d[x][i][0], y)
            # print(b)
            lis.append(b)
        e = be + lis
        lis_2 = []
        con = up_or_down(x, d, d2)
        lis_2.append(con)
        f = e + lis_2
        new_dic[x] = f
    return new_dic


'''function for list/dictionary for previous quarter'''
def making_dict_lq(names):
    d2 = {}
    for x in names:
        try:
            be = key_of_interest(x, data_previous_quarter, to_capture_previous_quarter)
            d2[x] = be
        except KeyError:
            pass
    return d2


'''function that places all information into the summary dashboard sheet'''
def placing_excel(d, d2):
    # loop through list/dictionary and place in correct place in DASHBOARD worksheet

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        print(project_name)
        if project_name in d:
            ws.cell(row=row_num, column=4).value = d[project_name][12][1]
            ws.cell(row=row_num, column=6).value = d[project_name][17][1]
            ws.cell(row=row_num, column=7).value = d[project_name][0][1]
            ws.cell(row=row_num, column=8).value = d[project_name][13][1]
            ws.cell(row=row_num, column=9).value = d[project_name][8][1]
            ws.cell(row=row_num, column=10).value = d[project_name][14][1]
            ws.cell(row=row_num, column=11).value = d[project_name][15][1]
            ws.cell(row=row_num, column=12).value = d[project_name][7][1]
            ws.cell(row=row_num, column=13).value = d[project_name][4][1]
            ws.cell(row=row_num, column=14).value = d[project_name][5][1]
            # dash_sheet.cell(row=row_num, column=27).value = d[project_name][9][1]
            # dash_sheet.cell(row=row_num, column=28).value = d[project_name][1][1]
            # dash_sheet.cell(row=row_num, column=33).value = d[project_name][2][1]
            # dash_sheet.cell(row=row_num, column=18).value = d[project_name][15][1]
            # dash_sheet.cell(row=row_num, column=23).value = d[project_name][16][1]
            # dash_sheet.cell(row=row_num, column=32).value = d[project_name][13][1]
            # dash_sheet.cell(row=row_num, column=37).value = d[project_name][14][1]

        if project_name == 'High Speed Rail Programme (HS2)':    #TODO should remove this hardcode
            ws.cell(row=row_num, column=13).value = 'often'
            ws.cell(row=row_num, column=14).value = 'often'

    for row_num in range(2, ws.max_row + 1):
        project_name = ws.cell(row=row_num, column=3).value
        if project_name in d2:
            ws.cell(row=row_num, column=5).value = d2[project_name][0][1]

    # Highlight cells that contain RAG text, with background and text the same colour. column E.
    ag_text = Font(color="00a5b700")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Green",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    ar_text = Font(color="00f97b31")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Red",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    red_text = Font(color="00fc2525")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Red",E1)))']
    ws.conditional_formatting.add('E1:E100', rule)

    green_text = Font(color="0017960c")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Green",e1)))']
    ws.conditional_formatting.add('E1:E100', rule)

    amber_text = Font(color="00fce553")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber",e1)))']
    ws.conditional_formatting.add('e1:e100', rule)

    # Highlight cells that contain RAG text, with background and black text columns G to L.
    ag_text = Font(color="000000")
    ag_fill = PatternFill(bgColor="00a5b700")
    dxf = DifferentialStyle(font=ag_text, fill=ag_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Green",G1)))']
    ws.conditional_formatting.add('G1:L100', rule)

    ar_text = Font(color="000000")
    ar_fill = PatternFill(bgColor="00f97b31")
    dxf = DifferentialStyle(font=ar_text, fill=ar_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber/Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber/Red",G1)))']
    ws.conditional_formatting.add('G1:L100', rule)

    red_text = Font(color="000000")
    red_fill = PatternFill(bgColor="00fc2525")
    dxf = DifferentialStyle(font=red_text, fill=red_fill)
    rule = Rule(type="containsText", operator="containsText", text="Red", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Red",G1)))']
    ws.conditional_formatting.add('G1:L100', rule)

    green_text = Font(color="000000")
    green_fill = PatternFill(bgColor="0017960c")
    dxf = DifferentialStyle(font=green_text, fill=green_fill)
    rule = Rule(type="containsText", operator="containsText", text="Green", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Green",G1)))']
    ws.conditional_formatting.add('G1:L100', rule)

    amber_text = Font(color="000000")
    amber_fill = PatternFill(bgColor="00fce553")
    dxf = DifferentialStyle(font=amber_text, fill=amber_fill)
    rule = Rule(type="containsText", operator="containsText", text="Amber", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("Amber",G1)))']
    ws.conditional_formatting.add('G1:L100', rule)

    # highlighting new projects
    red_text = Font(color="00fc2525")
    white_fill = PatternFill(bgColor="000000")
    dxf = DifferentialStyle(font=red_text, fill=white_fill)
    rule = Rule(type="containsText", operator="containsText", text="NEW", dxf=dxf)
    rule.formula = ['NOT(ISERROR(SEARCH("NEW",F1)))']
    ws.conditional_formatting.add('F1:F100', rule)

    # assign the icon set to a rule
    first = FormatObject(type='num', val=-1)
    second = FormatObject(type='num', val=0)
    third = FormatObject(type='num', val=1)
    iconset = IconSet(iconSet='3Arrows', cfvo=[first, second, third], percent=None, reverse=None)
    rule = Rule(type='iconSet', iconSet=iconset)
    ws.conditional_formatting.add('F1:F100', rule)

    # change text in last at next at BICC column
    for row_num in range(2, ws.max_row + 1):
        if ws.cell(row=row_num, column=13).value == '-2 weeks':
            ws.cell(row=row_num, column=13).value = 'Last BICC'
        if ws.cell(row=row_num, column=13).value == '2 weeks':
            ws.cell(row=row_num, column=13).value = 'Next BICC'
        if ws.cell(row=row_num, column=13).value == 'Today':
            ws.cell(row=row_num, column=13).value = 'This BICC'
        if ws.cell(row=row_num, column=14).value == '-2 weeks':
            ws.cell(row=row_num, column=14).value = 'Last BICC'
        if ws.cell(row=row_num, column=14).value == '2 weeks':
            ws.cell(row=row_num, column=14).value = 'Next BICC'
        if ws.cell(row=row_num, column=14).value == 'Today':
            ws.cell(row=row_num, column=14).value = 'This BICC'

            # highlight text in bold
    ft = Font(bold=True)
    for row_num in range(2, ws.max_row + 1):
        lis = ['This week', 'Next week', 'Last week', 'Two weeks',
               'Two weeks ago', 'This mth', 'Last mth', 'Next mth',
               '2 mths', '3 mths', '-2 mths', '-3 mths', '-2 weeks',
               'Today', 'Last BICC', 'Next BICC', 'This BICC',
               'Later this mth']
        if ws.cell(row=row_num, column=10).value in lis:
            ws.cell(row=row_num, column=10).font = ft
        if ws.cell(row=row_num, column=11).value in lis:
            ws.cell(row=row_num, column=11).font = ft
        if ws.cell(row=row_num, column=13).value in lis:
            ws.cell(row=row_num, column=13).font = ft
        if ws.cell(row=row_num, column=14).value in lis:
            ws.cell(row=row_num, column=14).font = ft
    return wb


'''keys of interest for current quarter'''
to_capture_current_quarter = ['Total Forecast', 'Departmental DCA', 'GMPP - IPA DCA', 'BICC approval point',
                              'Project Lifecycle Stage', 'Project MM20 Forecast - Actual',
                              'Project MM21 Forecast - Actual',
                              'SRO Finance confidence', 'Overall Resource DCA - Now',
                              'Overall Resource DCA - Future', 'SRO Benefits RAG', 'Last time at BICC', 'Next at BICC']

'''key of interest for previous quarter'''
to_capture_previous_quarter = ['Departmental DCA']

# 1) Provide file path to empty dashboard document
wb = load_workbook(
    'C:\\Users\\Standalone\\Will\\masters folder\\summary_dashboard_docs\\Q3_2018\\dashboard master_Q3_1819.xlsx')
ws = wb.active

# 2) Provide file path to master data sets
data_current_quarter = project_data_from_master(
    'C:\\Users\\Standalone\\Will\\masters folder\\core data\\merged_master_testing.xlsx')
data_previous_quarter = project_data_from_master(
    'C:\\Users\\Standalone\\Will\\masters folder\\core data\\master_2_2018.xlsx')

'''get list of project names'''
names = list(data_current_quarter.keys())
# names = ['A303 Amesbury to Berwick Down']   # can be useful for checking specific projects/the programme so leaving for now

'''creating mini dictionaries for the final command'''
d = making_dict(names)
d2 = making_dict_lq(names)
d3 = adding_con(d, d2)

'''command for running the programme'''
wb = placing_excel(d3, d2)

# 3) provide file path and specific name of output file.
wb.save(
    'C:\\Users\\Standalone\\Will\\masters folder\\summary_dashboard_docs\\Q3_2018\\testing_dashboard_Q3_2018_19.xlsx')
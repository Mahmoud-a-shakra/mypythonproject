import os
import json
import requests
import xlsxwriter
from datetime import datetime

from collections import OrderedDict
from xlsxwriter.utility import xl_col_to_name

ACCESS_TOKEN = ''
WORKBOOK = ''


def create_daysheet():
    worksheet = create_worksheet()

    races = get_races_today()
    write_meetings_to_sheet(worksheet=worksheet, races=races)
    favourites = write_all_races_to_sheet(worksheet=worksheet, races=races)

    write_fav_label_column(
        row=0,
        column=28,
        worksheet=worksheet,
        number_of_favourites=len(favourites)
    )

    write_races_list_to_sheet(
        row=0,
        column=29,
        worksheet=worksheet,
        races=races
    )

    write_protect_message(worksheet=worksheet)

    write_race_percentages(
        worksheet=worksheet,
        number_of_favourites=len(favourites)
    )

    WORKBOOK.close()


def create_workbook():
    today = datetime.today()
    return xlsxwriter.Workbook(
        'day-sheet-%s.xlsx' % today.strftime('%d-%m-%y')
    )


def create_worksheet():
    worksheet = WORKBOOK.add_worksheet()
    worksheet.protect('123')

    locked_format = get_locked_format()
    worksheet.set_column('A:XDF', None, locked_format)

    worksheet.set_default_row(12.6)
    worksheet.set_column('A:A', 4.11)
    worksheet.set_column('AA:AA', 3.11)
    worksheet.set_column('AB:AB', 5.56)
    cell_format = get_spacer_format()
    worksheet.set_column('AC:AC', 0.63, cell_format)
    worksheet.set_column('AF:AF', 4.56)
    worksheet.set_column('AG:AG', 14.11)
    worksheet.set_column('AH:AJ', 6.56)
    worksheet.set_column('B:K', 2.56)
    worksheet.set_column('L:L', 3.11)
    worksheet.set_column('P:P', 4.11)
    worksheet.set_column('Q:AA', 2.56)
    worksheet.set_column('M:M', 5.22)
    worksheet.set_column('N:N', 5.33)
    worksheet.set_column('O:O', 1.67)

    return worksheet


def get_spacer_format():
    cell_format = WORKBOOK.add_format()
    cell_format.set_bg_color('#000000')
    return cell_format


def get_locked_format():
    locked_format = WORKBOOK.add_format()
    locked_format.set_locked(True)
    return locked_format


def get_day_of_week_number():
    return int(datetime.today().strftime('%w')) + 1


def get_week_number():
    return int(datetime.today().strftime('%U')) + 1


def get_races_today():
    day = get_day_of_week_number()
    week = get_week_number()
    year = datetime.today().year

    url = ('xxx'
           'xxx/{year}-{week}-{day}'
           .format(
               **{'day': day, 'week': week, 'year': year}
           ))
    headers = {
        'x-api-key': 'xxx',
        'content-type': 'application/json; charset=UTF-8',
        'referer': 'xxx',
        'x-betfair-token': 'xxx',
        'authorization': ACCESS_TOKEN
    }
    response = requests.get(url, headers=headers)
    return response.json()


def get_racing_data(track_names):
    url = 'xx'
    headers = {
        'x-api-key': 'xxx',
        'content-type': 'application/json; charset=UTF-8',
        'referer': 'xxx',
        'authorization': ACCESS_TOKEN
    }
    day = get_day_of_week_number()
    week = get_week_number()
    data = {
        "bz02_ip_gt": 0,
        "bz04_ip_lt": 1002,
        "bz06_sp_gt": 1,
        "bz08_sp_lt": 1002,
        "bz10_wk_from": week,
        "bz12_wk_to": week,
        "bz14_ran_from": 1,
        "bz16_ran_to": 50,
        "bz18_distance": "",
        "bz20_ran": 0,
        "bz22_race_type": "",
        "bz24_fav": 0,
        "bz26_days_of_week": "%s" % day,
        "bz28_track": ', '.join(track_names),
        "bz30_ire_uk": "",
        "bz32_fav_sp1": 0,
        "bz33_fav_sp2": 0,
        "bz35_fav2_sp1": 0,
        "bz36_fav2_sp2": 0,
        "bz40_time_only1": 1.25,
        "bz41_time_only2": 1.958333,
        "bz43_fav_h": 0,
        "bz72_loser_winner": 0,
        "bz74_day": 0,
        "bz76_month": 0,
        "bz79_sp_gt": 0,
        "bz81_sp_lt": 0,
        "bz83_jmp_flat": 0,
        "years": "2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018",
        "bz38_hide_days": 0,
        "bz69_raceno": 0,
        "win_streak": 0,
        "bz65_ire_plus": 0,
        "bz49_racetype_2": "",
        "bz47_dis_2": ""
    }
    raw_response = requests.post(
        url,
        data=json.dumps(data).replace(' ', ''),
        headers=headers
    )

    return raw_response.json()


def split_races(races):
    track_names = list(
        set(
            [
                race['G_track']
                for race in races
            ]
        )
    )
    return {
        track_name: [
            race['N_time']
            for race in races
            if race['G_track'] == track_name
        ]
        for track_name in track_names
    }


def write_all_races_to_sheet(worksheet, races):
    races_split = split_races(races=races)
    track_names = races_split.keys()
    racing_data = get_racing_data(track_names=track_names)
    favourites = get_relevant_favourites(
        favourites=racing_data['yrs']['Fav']
    )
    write_meeting_to_sheet(
        worksheet=worksheet,
        favourites=favourites,
        row=0,
        start_column=16
    )
    set_this_year_all_formulas(
        row=2,
        worksheet=worksheet,
        number_of_favourites=len(favourites)
    )
    return favourites


def set_this_year_all_formulas(row, worksheet, number_of_favourites):
    winner_format = get_winner_format()
    for race_number in range(0, number_of_favourites):
        target_cell = 'AA{row}'.format(
            **{
                'row': row + race_number
            }
        )
        formula = 'COUNTIF(O:O,"{favourite_number}")'.format(
            **{
                'start_row': row,
                'end_row': row + number_of_favourites,
                'favourite_number': race_number + 1
            }
        )
        worksheet.write_formula(target_cell, formula, winner_format)


def write_meetings_to_sheet(worksheet, races):
    races_split = split_races(races=races)
    row = 0
    for track_name, times in races_split.items():
        racing_data = get_racing_data(track_names=[track_name])
        favourites = get_relevant_favourites(
            favourites=racing_data['yrs']['Fav']
        )

        if not bool(favourites):
            continue

        write_track_and_time_to_sheet(
            row=row,
            worksheet=worksheet,
            track_name=track_name,
            times=times
        )
        write_meeting_to_sheet(
            worksheet=worksheet,
            favourites=favourites,
            row=row,
            start_column=1
        )

        set_this_year_races_formula(
            row=row + 2,
            worksheet=worksheet,
            number_of_favourites=len(favourites),
            number_of_races=len(times)
        )

        row += len(favourites) + 4


def set_this_year_races_formula(row, worksheet, number_of_favourites, number_of_races):
    winner_format = get_winner_format()
    for favourite_number in range(0, number_of_favourites):
        target_cell = 'L{row}'.format(
            **{
                'row': row + favourite_number
            }
        )
        formula = 'COUNTIF(O{start_row}:O{end_row},"{favourite_number}")'.format(
            **{
                'start_row': row,
                'end_row': row + number_of_races,
                'favourite_number': favourite_number + 1
            }
        )
        worksheet.write_formula(target_cell, formula, winner_format)


def get_winner_format():
    winner_format = create_base_format()
    winner_format.set_align('center')
    winner_format.set_align('vcenter')
    winner_format.set_border(1)
    winner_format.set_border_color('#a6a6a6')
    return winner_format


def write_meeting_to_sheet(worksheet, favourites, row, start_column):

    write_race_headers(
        row=row,
        column=start_column,
        worksheet=worksheet,
        number_of_favourites=len(favourites.keys())
    )
    row += 1

    winner_format = get_winner_format()

    for favourite, years in favourites.items():
        column = start_column
        del years['all']
        for year, winners in years.items():
            if winners != 0:
                worksheet.write(row, column, winners, winner_format)
            else:
                worksheet.write(row, column, '', winner_format)
            column += 1
            worksheet.write(row, column, 0, winner_format)
        row += 1
    set_race_winners_conditional_format(
        row=row,
        start_column=start_column,
        number_of_favourites=len(favourites.keys()),
        worksheet=worksheet
    )
    write_totals_line(
        row=row,
        start_column=start_column,
        worksheet=worksheet,
        number_of_favourites=len(favourites.keys())
    )


def get_time_format():
    time_format = create_base_format()
    time_format.set_font_size(9)
    time_format.set_num_format('@')
    time_format.set_border(1)
    time_format.set_border_color('#000000')
    time_format.set_align('right')
    return time_format


def get_track_name_format():
    track_name_format = create_base_format()
    track_name_format.set_border(1)
    track_name_format.set_border_color('#000000')
    track_name_format.set_font_size(9)
    return track_name_format


def write_track_and_time_to_sheet(row, worksheet, track_name, times):
    column = 12
    set_winner_conditional_format(
        format_range='O1:O100',
        worksheet=worksheet
    )
    time_format = get_time_format()
    track_name_format = get_track_name_format()
    for time in times:
        row += 1
        worksheet.write(row, column, time, time_format)
        worksheet.write(row, column + 1, track_name, track_name_format)
        set_race_time_forumula(row=row + 1, column=column + 2, worksheet=worksheet)


def set_race_time_forumula(row, column, worksheet):
    track_name_format = get_track_name_format()
    start_column = xl_col_to_name(column)
    target_cell = '{column}{row}'.format(
        **{
            'column': start_column,
            'row': row
        }
    )

    formula = ('=IF(LEN(VLOOKUP({start_column}{start_row},AD:AF,3,FALSE))=0,"",'
               'VLOOKUP({start_column}{start_row},AD:AF,3,FALSE))').format(
        **{
            'start_column': xl_col_to_name(column - 2),
            'start_row': row
        }
    )
    worksheet.write_formula(target_cell, formula, track_name_format)


def set_winner_conditional_format(format_range, worksheet):
    fav_format = get_fav_format()
    outsider_format = get_outsider_format()
    blank_format = get_blank_format()
    worksheet.conditional_format(
        format_range,
        {
            'type': 'blanks',
            'format': blank_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': '<=',
            'value': 2,
            'format': fav_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': '>=',
            'value': 3,
            'format': outsider_format,
            'stop_if_true': True
        }
    )


def get_fav_format():
    fav_format = create_base_format()
    fav_format.set_pattern(1)
    fav_format.set_align('center')
    fav_format.set_align('vcenter')
    fav_format.set_bg_color('#FFFF00')
    return fav_format


def get_blank_format():
    blank_format = create_base_format()
    blank_format.set_align('center')
    blank_format.set_align('vcenter')
    return blank_format


def get_outsider_format():
    outsider_format = create_base_format()
    outsider_format.set_pattern(1)
    outsider_format.set_align('center')
    outsider_format.set_align('vcenter')
    outsider_format.set_font_color('#FFFFFF')
    outsider_format.set_bg_color('#203764')
    return outsider_format


def get_danger_format():
    danger_format = create_base_format()
    danger_format.set_pattern(1)
    danger_format.set_align('center')
    danger_format.set_align('vcenter')
    danger_format.set_font_color('#9c0000')
    danger_format.set_bg_color('#ffc7ce')
    return danger_format


def set_race_winners_conditional_format(row, start_column, number_of_favourites, worksheet):
    fav_format = get_fav_format()
    outsider_format = get_outsider_format()
    fav_format_range = '{start_column}{start_row}:{end_column}{end_row}'.format(
        **{
            'start_column': xl_col_to_name(start_column),
            'start_row': row - number_of_favourites + 1,
            'end_column': xl_col_to_name(start_column + 10),
            'end_row': row - number_of_favourites + 2
        }
    )
    outsider_format_range = '{start_column}{start_row}:{end_column}{end_row}'.format(
        **{
            'start_column': xl_col_to_name(start_column),
            'start_row': row - number_of_favourites + 3,
            'end_column': xl_col_to_name(start_column + 10),
            'end_row': row
        }
    )
    set_fav_conditional_format(
        format_range=fav_format_range,
        worksheet=worksheet,
        cell_format=fav_format
    )
    set_outsider_conditional_format(
        format_range=outsider_format_range,
        worksheet=worksheet,
        cell_format=outsider_format
    )


def set_fav_conditional_format(format_range, worksheet, cell_format):
    blank_format = get_blank_format()
    danger_format = get_danger_format()
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': '>=',
            'value': 1,
            'format': cell_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'blanks',
            'format': blank_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': danger_format,
            'stop_if_true': True
        }
    )


def set_outsider_conditional_format(format_range, worksheet, cell_format):
    blank_format = get_blank_format()
    danger_format = get_danger_format()
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': '>=',
            'value': 1,
            'format': cell_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'blanks',
            'format': blank_format,
            'stop_if_true': True
        }
    )
    worksheet.conditional_format(
        format_range,
        {
            'type': 'cell',
            'criteria': 'equal to',
            'value': 0,
            'format': danger_format,
            'stop_if_true': True
        }
    )


def get_total_format():
    total_format = create_base_format()
    total_format.set_pattern(1)
    total_format.set_align('center')
    total_format.set_align('vcenter')
    total_format.set_bg_color('#5E943C')
    total_format.set_font_color('#FFFFFF')
    return total_format


def write_totals_line(row, start_column, worksheet, number_of_favourites):
    for column_number in range(start_column, start_column + 11):
        column = xl_col_to_name(column_number)
        target_cell = '{column}{row}'.format(
            **{
                'column': column,
                'row': row + 1
            }
        )
        formula = '=SUM({column}{row_start}:{column}{row_end})'.format(
            **{
                'column': column,
                'row_start': row - number_of_favourites + 1,
                'row_end': row
            }
        )
        total_format = get_total_format()
        worksheet.write_formula(target_cell, formula, total_format)


def get_relevant_favourites(favourites):
    favourites = OrderedDict(favourites)
    to_delete = []
    del favourites['total']
    for favourite, years in reversed(favourites.items()):
        if favourites_not_empty(years=years):
            break
        to_delete.append(favourite)
    for favourite in to_delete:
        del favourites[favourite]
    return favourites


def write_race_headers(row, column, worksheet, number_of_favourites):
    cell_format = get_race_header_format()
    write_column_headers(
        row=row,
        column=column,
        worksheet=worksheet,
        number_of_favourites=number_of_favourites,
        cell_format=cell_format
    )
    worksheet.write(row, column - 1, 'Fav No', cell_format)
    write_row_headers(
        row=row,
        column=column,
        worksheet=worksheet,
        cell_format=cell_format
    )


def write_row_headers(row, column, worksheet, cell_format):
    for year in range(2009, 2020):
        if year != 2019:
            worksheet.write(row, column, str(year)[-2:], cell_format)
        else:
            worksheet.write(row, column, year, cell_format)
        column += 1


def write_column_headers(row, column, worksheet, number_of_favourites, cell_format):
    special_cases = {
        1: 'Fav',
        2: '2nd',
        3: '3rd',
        21: '21st',
        22: '22nd',
        23: '23rd'
    }
    for favourite in range(1, number_of_favourites + 1):
        row += 1
        if favourite in special_cases:
            worksheet.write(row, column - 1, special_cases[favourite], cell_format)
        else:
            worksheet.write(row, column - 1, '%sth' % favourite, cell_format)


def write_fav_label_column(row, column, worksheet, number_of_favourites):
    write_column_headers(
        row=row,
        column=column,
        worksheet=worksheet,
        number_of_favourites=number_of_favourites,
        cell_format=get_winner_format()
    )


def favourites_not_empty(years):
    for year, winners in years.items():
        if winners > 0:
            return True
    return False


def get_race_header_format():
    header_format = create_base_format()
    header_format.set_bold()
    header_format.set_pattern(1)
    header_format.set_align('center')
    header_format.set_align('vcenter')
    header_format.set_bg_color('#D9D9D9')
    header_format.set_font_color('#002060')
    header_format.set_border(1)
    header_format.set_border_color('#a6a6a6')
    header_format.set_num_format('@')
    return header_format


def create_base_format():
    workbook_format = WORKBOOK.add_format()
    workbook_format.set_font_name('Calibri')
    workbook_format.set_font_size(8)
    return workbook_format


def write_races_list_headers(row, column, worksheet):
    cell_format = get_race_header_format()
    cell_format.set_locked(False)
    worksheet.write(row, column, 'Time', cell_format)
    worksheet.write(row, column + 1, 'Track', cell_format)
    worksheet.write(row, column + 2, 'Winner', cell_format)
    worksheet.write(row, column + 3, 'Winner name', cell_format)


def write_races_list_to_sheet(row, column, worksheet, races):
    write_races_list_headers(row=row, column=column, worksheet=worksheet)
    cell_format = get_track_name_format()
    cell_format.set_locked(False)
    time_format = get_time_format()
    time_format.set_locked(False)

    for index, race in enumerate(races):
        worksheet.write(row + index + 1, column, race['N_time'], time_format)
        worksheet.write(row + index + 1, column + 1, race['G_track'], cell_format)
        if 'Winner' in race:
            winner = race['Winner']
        else:
            winner = ''
        if 'WinnerName' in race:
            winner_name = race['WinnerName']
        else:
            winner_name = ''
        worksheet.write(row + index + 1, column + 2, winner, cell_format)
        worksheet.write(row + index + 1, column + 3, winner_name, cell_format)

    start_column = xl_col_to_name(column + 2)
    start_row = row + 2
    end_row = row + len(races) + 1
    format_range = '{start_column}{start_row}:{end_column}{end_row}'.format(
        **{
            'start_column': start_column,
            'start_row': start_row,
            'end_column': start_column,
            'end_row': end_row
        }
    )
    set_winner_conditional_format(format_range=format_range, worksheet=worksheet)


def write_protect_message(worksheet):
    worksheet.write(0, 33, 'Password to unlock the sheet is 123')
    worksheet.write(2, 33, 'Only reason it is locked is to protect the formulas')


def write_race_percentages(worksheet, number_of_favourites):
    start_row = 4
    start_target_row = 2
    for favourite_number in range(0, number_of_favourites):
        worksheet.write(start_row + favourite_number, 33, favourite_number + 1)
        set_percentage_win_count_formula(
            start_row=start_row,
            favourite_number=favourite_number,
            worksheet=worksheet,
            start_target_row=start_target_row
        )
        set_percentage_win_formula(
            start_row=start_row,
            favourite_number=favourite_number,
            worksheet=worksheet,
            number_of_favourites=number_of_favourites
        )
    set_percentage_win_sum_formula(
        start_row=start_row,
        worksheet=worksheet,
        number_of_favourites=number_of_favourites
    )


def set_percentage_win_formula(start_row, favourite_number, worksheet, number_of_favourites):
    target_cell = 'AJ{row}'.format(
        **{
            'row': start_row + favourite_number + 1
        }
    )
    formula = '=IF($AI{total_row}=0,0.0,SUM(AI{row}/AI{total_row}))'.format(
        **{
            'row': start_row + favourite_number + 1,
            'total_row': start_row + number_of_favourites + 1
        }
    )
    cell_format = create_base_format()
    cell_format.set_num_format('0.00%')
    worksheet.write_formula(target_cell, formula, cell_format)


def set_percentage_win_sum_formula(start_row, worksheet, number_of_favourites):
    target_cell = 'AI{row}'.format(
        **{
            'row': start_row + number_of_favourites + 1
        }
    )
    formula = '=SUM(AI{start_row}:AI{end_row})'.format(
        **{
            'start_row': start_row + 1,
            'end_row': start_row + number_of_favourites - 1,
        }
    )
    cell_format = create_base_format()
    worksheet.write_formula(target_cell, formula, cell_format)


def set_percentage_win_count_formula(start_row, favourite_number, worksheet, start_target_row):
    target_cell = 'AI{row}'.format(
        **{
            'row': start_row + favourite_number + 1
        }
    )
    formula = '=AA{row}'.format(
        **{
            'row': start_target_row + favourite_number,
        }
    )
    cell_format = create_base_format()
    worksheet.write_formula(target_cell, formula, cell_format)


def login():
    url = 'xxx'
    headers = {
        'x-api-key': 'xxx',
        'content-type': 'application/json; charset=UTF-8',
        'referer': 'xxx'
    }
    data = '{"user":"xxx","pass":"%s"}' % os.environ.get("RSA_PASSWORD")
    raw_response = requests.post(url, data=data, headers=headers)
    response = raw_response.json()
    return response['accessToken']


if __name__ == '__main__':
    ACCESS_TOKEN = login()
    WORKBOOK = create_workbook()
    create_daysheet()
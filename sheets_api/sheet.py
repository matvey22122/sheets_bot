import re
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
CREDENTIALS_FILE = 'credentials.json'


def get_users():
    credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()

    data = sheet.values().get(spreadsheetId="1T28I53cIKO6Un9VnO7VMkOSnLQFQXwRkAbRD7ojLHSc",
                              range="Sheet1!B2:D15").execute()
    data = data.get('values', [])

    clean_data = []
    for i in data:
        ar = i
        if len(ar) != 3:
            ar.append("")
        clean_data.append(ar)

    for i in range(len(clean_data)):
        clean_data[i][0] = clean_data[i][0].split("/")[5]
        if clean_data[i][2]:
            clean_data[i][2] = clean_data[i][2].split("/")[5]

    return clean_data


def _get_table_name():
    days_to_subtract = datetime.now().weekday()

    if days_to_subtract >= 5:
        days_to_subtract = 7 - days_to_subtract
        date = datetime.now() + timedelta(days=days_to_subtract)
    else:
        date = datetime.now() - timedelta(days=days_to_subtract)

    # date = date - timedelta(days=7)  # TODO: remove!!!!
    table_name = f"{date.day} {date.strftime('%B')} {date.year}"
    table_inner = date.strftime('%m/%d/%Y')
    table_past = f"{(date - timedelta(days=7)).day} {(date - timedelta(days=7)).strftime('%B')} {(date - timedelta(days=7)).year}"
    date += timedelta(days=7)
    table_future_name = f"{date.day} {date.strftime('%B')}"
    return [table_name, date, table_future_name, table_inner, table_past]


def _get_table_name2():
    days_to_subtract = datetime.now().weekday()

    if days_to_subtract >= 5:
        days_to_subtract = 7 - days_to_subtract
        date = datetime.now() + timedelta(days=days_to_subtract)
    else:
        date = datetime.now() - timedelta(days=days_to_subtract)

    date += timedelta(days=7)
    # date = date - timedelta(days=7)  # TODO: remove!!!!
    table_name = f"{date.day} {date.strftime('%B')} {date.year}"
    table_inner = date.strftime('%m/%d/%Y')
    table_past = f"{(date - timedelta(days=7)).day} {(date - timedelta(days=7)).strftime('%B')} {(date - timedelta(days=7)).year}"
    date += timedelta(days=7)
    table_future_name = f"{date.day} {date.strftime('%B')}"
    return [table_name, date, table_future_name, table_inner, table_past]


def _get_days_to_catch():
    days_to_catch = []

    days_to_subtract = datetime.now().weekday()

    if days_to_subtract >= 5:
        days_to_subtract = 7 - days_to_subtract
        date = datetime.now() + timedelta(days=days_to_subtract)
    else:
        date = datetime.now() - timedelta(days=days_to_subtract)

    date = date - timedelta(days=2)
    # date = date - timedelta(days=7)  # TODO: remove!!!!
    days_to_catch.append([date.day, date.month, date.year])
    for i in range(6):
        date += timedelta(days=1)
        days_to_catch.append([date.day, date.month, date.year])

    return days_to_catch


class Sheet:
    credentials = service_account.Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)
    sheet = service.spreadsheets()

    days_to_catch = _get_days_to_catch()
    table_name, table_time, table_future_name, table_inner, table_past = _get_table_name()
    table_name2, table_time2, table_future_name2, table_inner2, table_past2 = _get_table_name2()

    # print(days_to_catch, table_name, table_time)

    def __init__(self, from_table, to_table, work_hours):
        self.time_itervals = [["A", "B"], ["F", "G"], ["K", "L"], ["P", "Q"], ["U", "V"], ["Z", "AA"], ["AE", "AF"]]
        self.from_table = from_table
        self.to_table = to_table
        self.work_hours = work_hours

    def create_sheet(self):
        results = self.sheet.sheets().copyTo(spreadsheetId="1e28J_83AL8hrBtvLuBarlZnjAbEJmXYQOgOMVGyxWfk",
                                             sheetId="263768700",
                                             body={'destination_spreadsheet_id': self.to_table}).execute()

        self.sheet.batchUpdate(spreadsheetId=self.to_table, body={
            'requests': [
                {"updateSheetProperties": {
                    "properties": {
                        "sheetId": results['sheetId'],
                        "title": self.table_name,
                    },
                    "fields": "title",
                }}
            ]
        }).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name + f"!K3:K3",
            valueInputOption="USER_ENTERED",
            body={
                "values": [[self.table_inner]]}).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name + f"!B1:B1",
            valueInputOption="USER_ENTERED",
            body={
                "values": [[f"='{self.table_past}'!L1"]]}).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name + f"!J1:J1",
            valueInputOption="USER_ENTERED",
            body={"values": [[f"=СЧИТАТЬПУСТОТЫ(E3,AI3,AD3,Y3,T3,O3,J3)*12*{self.work_hours.split(':')[0]}"]]}).execute()

    def create_sheet2(self):
        results = self.sheet.sheets().copyTo(spreadsheetId="1e28J_83AL8hrBtvLuBarlZnjAbEJmXYQOgOMVGyxWfk",
                                             sheetId="263768700",
                                             body={'destination_spreadsheet_id': self.to_table}).execute()

        self.sheet.batchUpdate(spreadsheetId=self.to_table, body={
            'requests': [
                {"updateSheetProperties": {
                    "properties": {
                        "sheetId": results['sheetId'],
                        "title": self.table_name2,
                    },
                    "fields": "title",
                }}
            ]
        }).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name2 + f"!K3:K3",
            valueInputOption="USER_ENTERED",
            body={
                "values": [[self.table_inner2]]}).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name2 + f"!B1:B1",
            valueInputOption="USER_ENTERED",
            body={
                "values": [[f"='{self.table_past2}'!L1"]]}).execute()

        self.sheet.values().update(
            spreadsheetId=self.to_table,
            range=self.table_name2 + f"!J1:J1",
            valueInputOption="USER_ENTERED",
            body={"values": [[f"=СЧИТАТЬПУСТОТЫ(E3,AI3,AD3,Y3,T3,O3,J3)*12*{self.work_hours.split(':')[0]}"]]}).execute()

    def transport_data(self):
        try:
            row = \
                self.sheet.values().get(spreadsheetId=self.to_table, range=f'{self.table_name}!A3:AI3').execute().get(
                    'values',
                    [])[0]
        except Exception as e:
            self.create_sheet()
            row = \
                self.sheet.values().get(spreadsheetId=self.to_table, range=f'{self.table_name}!A3:AI3').execute().get(
                    'values',
                    [])[0]

        try:
            row = \
                self.sheet.values().get(spreadsheetId=self.to_table, range=f'{self.table_name2}!A3:AI3').execute().get(
                    'values',
                    [])[0]
        except Exception as e:
            self.create_sheet2()

        if self.from_table == "":
            return

        data = self.sheet.values().get(spreadsheetId=self.from_table, range="Sheet1!J500:M100000").execute()
        data = self._prepare_data(data.get('values', []))
        # print(data)
        interval_for_comments = ["E", "J", "O", "T", "Y", "AD", "AI"]
        for i in range(7):
            results = self.sheet.values().clear(
                spreadsheetId=self.to_table,
                range=f'{self.table_name}!{self.time_itervals[i][0]}4:{self.time_itervals[i][1]}16').execute()

            results = self.sheet.values().clear(
                spreadsheetId=self.to_table,
                range=f'{self.table_name}!{interval_for_comments[i]}4:{interval_for_comments[i]}16').execute()

            number = str(self.days_to_catch[i][0])
            if number not in data:
                continue

            records = []
            for j in data[number]:
                records.append(j[:2])
            body = {
                "values": records,
            }

            results = self.sheet.values().update(
                spreadsheetId=self.to_table,
                range=f'{self.table_name}!{self.time_itervals[i][0]}4:{self.time_itervals[i][1]}{4 + len(data[number]) - 1}',
                valueInputOption="RAW",
                body=body).execute()

            for j in range(len(data[number])):
                if len(data[number][j]) == 3:
                    results = self.sheet.values().update(
                        spreadsheetId=self.to_table,
                        range=f'{self.table_name}!{interval_for_comments[i]}{4 + j}:{interval_for_comments[i]}{4 + j}',
                        valueInputOption="RAW",
                        body={"values": [[data[number][j][2]]]}).execute()

    def _prepare_data(self, data):
        clean_data = {}

        for row in data:
            # check if 'record' exist
            if row[0] == "":
                continue

            #если не в одном дне
            if row[0] != row[2]:
                if [datetime.strptime(row[0], '%d/%m/%Y').day, datetime.strptime(row[0], '%d/%m/%Y').month,
                    datetime.strptime(row[0], '%d/%m/%Y').year] in self.days_to_catch:
                    if str(datetime.strptime(row[0], '%d/%m/%Y').day) not in clean_data:
                        clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)] = []
                    if [row[1], "23:59"] not in clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)]:
                        clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)].append([row[1], "23:59"])

                if [datetime.strptime(row[2], '%d/%m/%Y').day, datetime.strptime(row[2], '%d/%m/%Y').month,
                    datetime.strptime(row[2], '%d/%m/%Y').year] in self.days_to_catch:
                    if str(datetime.strptime(row[2], '%d/%m/%Y').day) not in clean_data:
                        clean_data[str(datetime.strptime(row[2], '%d/%m/%Y').day)] = []
                    if ["00:00", row[3]] not in clean_data[str(datetime.strptime(row[2], '%d/%m/%Y').day)]:
                        clean_data[str(datetime.strptime(row[2], '%d/%m/%Y').day)].append(["00:00", row[3]])
            elif [datetime.strptime(row[0], '%d/%m/%Y').day, datetime.strptime(row[0], '%d/%m/%Y').month,
                  datetime.strptime(row[0], '%d/%m/%Y').year] in self.days_to_catch:
                if str(datetime.strptime(row[0], '%d/%m/%Y').day) not in clean_data:
                    clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)] = []
                if [row[1], row[3]] not in clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)]:
                    clean_data[str(datetime.strptime(row[0], '%d/%m/%Y').day)].append([row[1], row[3]])

        for d in clean_data.keys():
            for i in range(len(clean_data[d])):
                if len(clean_data[d][i][0]) < 5:
                    clean_data[d][i][0] = "0" + clean_data[d][i][0]
                if len(clean_data[d][i][1]) < 5:
                    clean_data[d][i][1] = "0" + clean_data[d][i][1]

        for d in clean_data.keys():
            for i in range(len(clean_data[d])):
                li = datetime.strptime(clean_data[d][i][0], "%H:%M")
                ri = datetime.strptime(clean_data[d][i][1], "%H:%M")
                if ri < li:
                    continue
                for j in range(len(clean_data[d])):
                    if i == j:
                        continue
                    lj = datetime.strptime(clean_data[d][j][0], "%H:%M")
                    rj = datetime.strptime(clean_data[d][j][1], "%H:%M")

                    if rj < lj:
                        continue

                    if li <= lj <= ri or li <= rj <= ri:
                        if li.strftime("%H:%M") == "00:00":
                            clean_data[d][i] = [li.strftime("%H:%M"), li.strftime("%H:%M"), f"фича {clean_data[d][i][1]}"]
                        else:
                            clean_data[d][i] = [li.strftime("%H:%M"), (li - timedelta(seconds=1)).strftime("%H:%M"), f"!!! {clean_data[d][i][1]}"]

                        if lj.strftime("%H:%M") == "00:00":
                            clean_data[d][j] = [lj.strftime("%H:%M"), lj.strftime("%H:%M"), f"фича {clean_data[d][j][1]}"]
                        else:
                            clean_data[d][j] = [lj.strftime("%H:%M"), (lj - timedelta(seconds=1)).strftime("%H:%M"), f"!!! {clean_data[d][j][1]}"]

            clean_data[d].sort()
        return clean_data

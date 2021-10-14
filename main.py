from sheets_api.sheet import Sheet, get_users


def main():
    users = get_users()
    for user in users:
        sheet = Sheet(from_table=user[2], to_table=user[0], work_hours=f'{user[1]}:00')
        sheet.transport_data()


if __name__ == '__main__':
    main()

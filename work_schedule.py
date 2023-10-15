import openpyxl
from datetime import datetime, time, timedelta, date
from activity import get_user_activity_dict

def get_schedule_dict():
    #work_schedule_dict = {"id": {"work_schedule": [(date_start,start_hour,finish_hour)...]}}
    work_schedule_dict = {}


    workbook = openpyxl.load_workbook("Grafik Testowy.xlsx")
    worksheet = workbook["Wrzesień"]

    title_list = [cell.value for cell in worksheet[2]]

    title_dict = {
        "id": title_list.index("Nr karty"),
        "department": title_list.index("Dział"),
        "firm": title_list.index("Firma"),
        "position": title_list.index("Stanowisko"),
        "boss": title_list.index("Przełożony")

    }

    start_schedule_index = next((i for i, item in enumerate(title_list) if isinstance(item, datetime)), None)


    for row in worksheet.iter_rows(min_row=3, values_only=True):
        work_schedule = []
        card_id = row[title_dict["id"]]
        department = row[title_dict["department"]]
        firm = row[title_dict["firm"]]
        position = row[title_dict["position"]]
        boss = row[title_dict["boss"]]



        if card_id not in work_schedule_dict:
            work_schedule_dict[card_id] = {}
        row = row[start_schedule_index:]
        for index, value in enumerate(row[::2]):
            work_schedule.append((title_list[index*2+start_schedule_index] ,value, row[index*2+1]))
        work_schedule_dict[card_id]["work_schedule"] = work_schedule
        work_schedule_dict[card_id]["department"] = department
        work_schedule_dict[card_id]["firm"] = firm
        work_schedule_dict[card_id]["position"] = position
        work_schedule_dict[card_id]["boss"] = boss


    return work_schedule_dict


def get_actual_start_time(activity_list, shadule_start_time):
    """ return actuale datetime going to work if worker late return status 'yes' """
    filtered_times = [time for time in activity_list if time <= shadule_start_time]
    status = None

    if filtered_times:
        actual_start_time = max(filtered_times, key=lambda x: x)
        return actual_start_time, status

    else:
        actual_start_time = min(activity_list, key=lambda x: x)
        if actual_start_time.day == shadule_start_time.day:
            status = "yes"
            return actual_start_time, status
        else:
            print("брак даних")
            return None, status


def get_actual_finish_time(activity_list, shadule_finish_time):
    """ return actuale datetime going out work if worker erlier return status 'yes' """
    filtered_times = [time for time in activity_list if time >= shadule_finish_time]
    status = None

    if filtered_times:
        actual_finish_time = min(filtered_times, key=lambda x: x)
        return actual_finish_time, status

    else:
        actual_finish_time = max(activity_list, key=lambda x: x)
        if actual_finish_time.day == shadule_finish_time.day:
            status = "yes"
            return actual_finish_time, status
        else:
            print("брак даних")
            return None, status


def get_total_breakfast(user_activity):
    total_time = timedelta()  # Ініціалізуємо total_time як timedelta
    user_activity = user_activity[1:-1]
    exit = None  # Встановлюємо exit як None, щоб почати

    if user_activity:
        for datetime_, action in user_activity:
            if action == "exit":
                exit = datetime_
            elif action == "entrance" and exit is not None:
                total_time += datetime_ - exit
    return total_time


def analize_breakfast(start_datetime, finish_datetime, length_lunch):
    time_difference = finish_datetime - start_datetime
    work_hours = time_difference.total_seconds() / 3600

    if work_hours <= 8 and length_lunch > 30:
        return length_lunch
    if work_hours > 8 and length_lunch > 60:
        return length_lunch
    return ""

def analize_user(user_dict:dict, id_card:int):
    """return remark row if user have some discrepancies with shadule"""
    MAX_TIME_PERIOD = timedelta(hours=8)

    shadule = user_dict["work_schedule"]
    activity = user_dict["activity"]
    for worker_information in shadule:
        PRINT = False
        activity_list = []
        entrance_list = []
        exit_list = [] 
        total_breacfest = timedelta()
        remark_row = [
            "", 
            user_dict["department"], 
            user_dict["full_name"], 
            id_card, user_dict["firm"], 
            user_dict["position"], 
            user_dict["boss"], 
            "", "", "", ""
            ]

        date_start, start_hour,finish_hour = worker_information
        date_finish = date_start
        remark_row[0] = date_start


        if isinstance(start_hour, str):
            continue

        datetime_start_control = datetime.combine(date_start.date(), start_hour)
        datetime_finish_control = datetime.combine(date_start.date(), finish_hour)

        #визначення денна чи нічна зміна
        if finish_hour < start_hour:
            date_finish += timedelta(days=1) + MAX_TIME_PERIOD
            datetime_finish_control += timedelta(days=1)
        else:
            date_finish += timedelta(hours=23, minutes=59, seconds=59)

        # створення списку дій працівника за період ло крнтролі
        for date_time, event in activity:

            if date_time >= date_start and date_time <= date_finish:
                if event == "entrance":
                    entrance_list.append(date_time)
                    activity_list.append((date_time, "entrance"))
                if event == "exit":
                    exit_list.append(date_time)
                    activity_list.append((date_time, "exit"))

        if entrance_list:
            actual_start_time, status = get_actual_start_time(entrance_list, datetime_start_control)
            if actual_start_time:
                for index, time_action in enumerate(activity_list):
                    if time_action[0] == actual_start_time:
                        activity_list = activity_list[index:]
            if status:
                PRINT = True
                remark_row[7] = actual_start_time

        if exit_list:
            actual_finish_time, status = get_actual_finish_time([activity[0] for activity in activity_list], datetime_finish_control)
            if actual_finish_time:
                for index, time_action in enumerate(activity_list):
                    if time_action[0] == actual_finish_time:
                        activity_list = activity_list[:index+1]
            if status:
                PRINT = True
                remark_row[8] = actual_finish_time

        if activity_list:
            if activity_list[0][1] != "entrance":
                remark_row[10] += "entrance error"
                PRINT = True

            if activity_list[-1][1] != "exit":
                remark_row[10] += "exit error"
                PRINT = True


        if len(activity_list) >= 4:
            total_breacfest = get_total_breakfast(activity_list).total_seconds() // 60
            start_datetime = datetime.combine(datetime_start_control, start_hour)
            finish_datetime = datetime.combine(datetime_finish_control, finish_hour)
            breacfest = analize_breakfast(start_datetime, finish_datetime, total_breacfest)
            if breacfest:
                remark_row[6] = breacfest
                PRINT = True
        else:
            remark_row[6] = ""

        if PRINT:
            print(remark_row)
            print(" ")


def get_report():
    shadule = get_schedule_dict()
    workers_activity_dickt = get_user_activity_dict(shadule.keys())
    #workers_activity_dickt = get_user_activity_dict([2107416781])

    for worker in workers_activity_dickt:
        workers_activity_dickt[worker]["work_schedule"] = shadule[worker]["work_schedule"]
        workers_activity_dickt[worker]["department"] = shadule[worker]["department"]
        workers_activity_dickt[worker]["firm"] = shadule[worker]["firm"]
        workers_activity_dickt[worker]["position"] = shadule[worker]["position"]
        workers_activity_dickt[worker]["boss"] = shadule[worker]["boss"]
        

    # analize_user(workers_activity_dickt[2107435501], 2107435501)
    for user in workers_activity_dickt:
        analize_user(workers_activity_dickt[user], id_card=user)

get_report()
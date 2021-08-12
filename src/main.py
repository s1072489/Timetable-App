# Import dependencies
import threading
import schedule
import datetime
import pandas
import time
import os
from pystray import *
from PIL import Image

# System Tray
def exit():
	os._exit(0)

# Calculate
def now():
	now_run(systemTray=systemTray)
def now_run(systemTray):

	# Get current time in 24 format
	rn = datetime.datetime.now()
	hour = rn.hour
	minute = rn.minute
	rn = int(str(hour) + "00") + minute

	# Checks for the period that is directly after the said time
	for i in timetable.loc[:, "Time"]:
		if rn < i:
			period = i
			break
		else:
			pass

	# We only need week and day
	year, week, day = datetime.date.today().isocalendar()
	
	# Check if it is week 1 or week 2
	if week % 2 == 0:
		pass
	else: # If it is week 2, add 5 days as it matches with the excel timetable
		day += 5

	# Get the row which has the same time as the next period
	period_week = timetable.iloc[timetable.loc[timetable["Time"] == period].index].to_dict()

	# Get needed information
	info = list(period_week[f"Day {day}"].values())[0].split("|")
	subject = info[0] # The first thing in the column of the day (because the timetable is underneath)
	if subject != "None":
		period = list(period_week["Period"].values())[0] # Get the period number from the period column
		room = info[1]
		teacher = info[2]
		trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
			title = f"{subject} is starting next.",
			message = f"Room: {room}\nTeacher: {teacher}\nPeriod {period}"
		)
	else:
		pass

	time.sleep(30*60)

# Current
def last():
	last_run(systemTray=systemTray)
def last_run(systemTray):

	# Get current time in 24 format
	time = datetime.datetime.now()
	hour = time.hour
	minute = time.minute
	time = int(str(hour) + "00") + minute

	# Checks for the period that is directly after the said time
	b = None
	for i in timetable.loc[:, "Time"]:
		if time < i:
			period = b
			break
		else:
			b = i

	try:
		if period == None:
			trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
				title = f"There is no class?",
				message = f"Variable period not assigned"
			)
		else:
			# We only need week and day
			year, week, day = datetime.date.today().isocalendar()
			
			# Check if it is week 1 or week 2
			if week % 2 == 0:
				pass
			else: # If it is week 2, add 5 days as it matches with the excel timetable
				day += 5

			# Get the row which has the same time as the next period
			period_week = timetable.iloc[timetable.loc[timetable["Time"] == period].index].to_dict()

			# Get needed information
			info = list(period_week[f"Day {day}"].values())[0].split("|")
			subject = info[0] # The first thing in the column of the day (because the timetable is underneath)
			if subject != "None":
				period = list(period_week["Period"].values())[0] # Get the period number from the period column
				room = info[1]
				teacher = info[2]
				trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
					title = f"{subject} has started",
					message = f"Room: {room}\nTeacher: {teacher}\nPeriod {period}"
				)
			else:
				pass
	except:
		trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
				title = f"There is no class?",
				message = f"Variable period not assigned"
			)


# Manual
def manual():
	manual_run(systemTray=systemTray)
def manual_run(systemTray):

	# Get current time in 24 format
	time = datetime.datetime.now()
	hour = time.hour
	minute = time.minute
	time = int(str(hour) + "00") + minute

	# Checks for the period that is directly after the said time
	for i in timetable.loc[:, "Time"]:
		if time < i:
			period = i
			break
		else:
			pass

	try:
		period == 1
	except:
		trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
			title = f"There is no next class?",
			message = f"Variable period not assigned"
		)


	# We only need week and day
	year, week, day = datetime.date.today().isocalendar()
	
	# Check if it is week 1 or week 2
	if week % 2 == 0:
		pass
	else: # If it is week 2, add 5 days as it matches with the excel timetable
		day += 5

	# Get the row which has the same time as the next period
	period_week = timetable.iloc[timetable.loc[timetable["Time"] == period].index].to_dict()

	# Get needed information
	info = list(period_week[f"Day {day}"].values())[0].split("|")
	subject = info[0] # The first thing in the column of the day (because the timetable is underneath)
	if subject != "None":
		period = list(period_week["Period"].values())[0] # Get the period number from the period column
		room = info[1]
		teacher = info[2]
		trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
			title = f"{subject} is starting next.",
			message = f"Room: {room}\nTeacher: {teacher}\nPeriod {period}"
		)
	else:
		period_week = timetable.iloc[timetable.loc[timetable["Time"] == period].index - 1].to_dict()
		info = list(period_week[f"Day {day}"].values())[0].split("|")
		subject = info[0]
		period = list(period_week["Period"].values())[0] # Get the period number from the period column
		room = info[1]
		teacher = info[2]
		trayItem.notify( # TrayItem will be the same because it was passed into systemTray() along with notifyLoop()
			title = f"{subject} is a double.",
			message = f"Room: {room}\nTeacher: {teacher}\nPeriod {period}"
		)


timetable = pandas.read_excel("main.xlsx", sheet_name='Sheet1')

image = Image.open("Bell.png")
menu = Menu(MenuItem('Now', last), MenuItem('Next', manual), MenuItem('Exit', exit))
trayItem = Icon('Timetable', image, menu=menu)

for i in ["08:40","09:30", "10:15", "11:30", "12:20", "14:00", "14:50"]:
	schedule.every().monday.at(i).do(now)
	schedule.every().tuesday.at(i).do(now)
	schedule.every().wednesday.at(i).do(now)
	schedule.every().thursday.at(i).do(now)
	schedule.every().friday.at(i).do(now)

def systemTray(trayItem):
	trayItem.run()

# Wait for lesson
def notifyLoop(trayItem):
	while True:
		schedule.run_pending()
		time.sleep(1)

# Run required functions
notifs = threading.Thread(target=notifyLoop, args=(trayItem, ))
notifs.start()

systemTray(trayItem)

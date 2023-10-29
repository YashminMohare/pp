import subprocess
import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import smtplib
import ctypes
import time
import requests
import shutil
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

# its used for set voice of assistant
 
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

#  its start of code when os take cmd from user
 
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    
#  its wish Automatic time detect and wish fuction

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Sir !")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Sir !")

	else:
		speak("Good Evening Sir !")

	assname =("i am jarvis")
	speak("I am your Aifriend")
	speak(assname)
 
# username function for understand ur name .
 	
def username():
	speak("What should i call you Sir")
	uname = takeCommand()
	speak("Welcome")
	speak(uname)
	speak("sir ")
	columns = shutil.get_terminal_size().columns
	
	print("******************************************************".center(columns))
	print("Welcome in my AI",uname.center(columns))
	speak("welcom i my AI world")
	print("*******************************************************".center(columns))


# this function take data from user using voice recongnizer

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		r.pause_threshold = 1
		audio = r.listen(source)
  
# if user cmd properly recongnize then its print otherwise exception is run

	try:
		print("Recognizing...")
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

# its exception if cmd can't properly recognize 

	except Exception as e:
		print(e)
		print("Unable to Recognize your voice.")
		return "None"
	
	return query

if __name__ == '__main__':
    	clear = lambda: os.system('cls')
clear()
wishMe()
username()
		
speak("How can i Help you, baby")

# its loop for take continue cmd 

while True:
		
		#its comvert cmd to lowercase 

		query = takeCommand().lower()
		
		# for serach anything in wikipedia they show and also speak the result

		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")

			# its print and speak  only 3 line 

			results = wikipedia.summary(query, sentences = 3)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		#  for open Youtube

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			webbrowser.open("youtube.com")

		#  for open google

		elif 'open google' in query:
			speak("Here you go to Google\n")
			webbrowser.open("google.com")

		#  for open stack overflow 

		elif 'open stack overflow' in query:
			speak("Here you go to Stack Over flow.Happy coding")
			webbrowser.open("stackoverflow.com")

		#  play song if available in directory

		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			music_dir = 'E:\\dj best song'
			songs = os.listdir(music_dir)
			print(songs)
			random = os.startfile(os.path.join(music_dir, songs[1]))

		# to show and speak current time

		elif 'the time' in query:
			strTime1 =str(datetime.datetime.today()).split()[0]
			strTime2 = datetime.datetime.now().strftime(" %H:%M:%S ")
			speak("Sir the date is ")
			speak(strTime1)
			print(strTime1)
			speak("Sir the time is ")
			speak(strTime2)
			print(strTime2)

		#  normal conversation

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "what is your name" in query:
			speak("My friends call me")
			speak("jarvis")
			print("My friends call me")

		#  for exits from loop 

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query:
			speak("I have been created by yashmin.")
								
		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		elif "why you came to world" in query:
			speak("Thanks to yashmin. further It's a secret")

		#  for open ppt

		elif 'PowerPoint presentation' in query:
			speak("opening Power Point presentation")
			power = r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office PowerPoint 2007.lnk"
			os.startfile(power)

		#  open ms word

		elif 'ms word' in query:
			speak("opening ms word ")
			power = r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office Word 2007.lnk"
			os.startfile(power)

		#  open ms excel 

		elif 'ms excel' in query:
			speak("opening ms excel")
			power = r"C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office\\Microsoft Office Excel 2007.lnk"
			os.startfile(power)

		elif 'is love' in query:
			speak("It is 7th sense that destroy all other senses")

		elif "who are you" in query:
			speak("I am your virtual assistant created by yashmin mohare")

		elif 'reason for you' in query:
			speak("I was created as a Minor project by Mister yashmin mohare ")

		# change the background of your system

		elif 'change background' in query:
			ctypes.windll.user32.SystemParametersInfoW(20,
													0,
													"Location of wallpaper",
													0)
			speak("Background changed successfully")

		#  for lock window 

		elif 'lock window' in query:
				speak("locking the device")
				ctypes.windll.user32.LockWorkStation()

		# for shutdown your system

		elif 'shutdown system' in query:
				speak("Hold On a Sec ! Your system is on its way to shut down")
				subprocess.call('shutdown / p /f')
					
		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop jarvis from listening commands")
			a = int(takeCommand())
			time.sleep(a)
			print(a)
				
		elif "restart" in query:
			subprocess.call(["shutdown", "/r"])
			
		elif "hibernate" in query or "sleep" in query:
			speak("Hibernating")
			subprocess.call("shutdown / h")

		elif "log off" in query or "sign out" in query:
			speak("Make sure all the application are closed before sign-out")
			time.sleep(5)
			subprocess.call(["shutdown", "/l"])

		#  for create a note with voice cmd

		elif "write a note" in query:
			speak("What should i write, sir")
			note = takeCommand()
			file = open('jarvis.txt', 'w')
			speak("Sir, Should i include date and time")
			snfm = takeCommand()

			if 'yes' in snfm or 'sure' in snfm:
				strTime = datetime.datetime.now().strftime(" %H:%M:%S")
				file.write(strTime)
				file.write(" :- ")
				file.write(note)

			else:
				file.write(note)

		#  its speak ur privious notes 

		elif "give note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r")
			print(file.read())
			speak(file.read(6))
			
		elif "jarvis" in query:
			speak("yes sir")
		  				
        

		
		
		
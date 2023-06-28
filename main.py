import os
import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import pywhatkit
speaker = win32com.client.Dispatch("SAPI.SpVoice")
import smtplib
from email.message import EmailMessage


# Functions

# Taking Voice Command Function
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold=0.5
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f" User said {query}")
            return query
        except Exception as e:
            speaker.Speak("Sorry didn't heard you")
            print("Sorry didn't heard you ")



# Sending Email Functions
def send_email(subject,message,receiver):
    email_address = '###################'  # Enter your Email Id
    password = '' # Enter your Password


    msg = EmailMessage()
    msg.set_content(message)
    msg['Subject'] = subject
    msg['From'] = email_address
    msg['To'] = receiver

    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        smtp.login(email_address, password)
        smtp.send_message(msg)


speaker.Speak("Hello I am Jarvis")
print("Hello I am AI BOT")




# List
Email={
    "me":"###############"
}
contacts = {

    "Me": "+91#########"
}
sites = {
        "Youtube": "www.youtube.com", "Google": "www.google.com", "Wikipedia": "www.wikipedia.com"
    }
apps = {
        "Notepad": "notepad.exe",
        "Teams": "C:\\Users\\vyoms\\AppData\\Local\\Microsoft\\Teams\\current\\Teams.exe",
        "Python": "C:\\Program Files\\JetBrains\\PyCharm 2023.1.3\\bin\\pycharm64.exe",
        "Java":"C:\\Program Files\\JetBrains\\IntelliJ IDEA Community Edition 2021.3.2\\bin\\idea64.exe",
        "vs code": "C:\\Users\\vyoms\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe",
        "excel":"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
        "powerpoint":"C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
        "word":"C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
        # Enter application name with its path
        }


folder={
        "download":"C:\\Users\\users\\Downloads",
        # Enter folder name with its path
    }

while 1:

    print("Listening....")

    text = takecommand()
    if text is None:
        speaker.Speak("Waiting For your Command")
        continue


    speaker.Speak(text)



    # Open for Opening website, Application and Folder
    if "Open".lower() in text.lower():
        if "website" in text.lower():
            for site in sites:
                if f"Open website {site}".lower() in text.lower():
                    speaker.Speak(f"Opening {site}")
                    webbrowser.open(sites[site])
        elif "folder" in text.lower():
            for fold in folder:
                if f"Open {fold}".lower() in text.lower():
                    speaker.Speak(f"Opening {fold} Folder")
                    os.startfile(folder[fold])

        else:
            for app in apps:
                if f"Open {app}".lower() in text.lower():
                   speaker.Speak(f"Opening {app}")
                   os.startfile(apps[app])


    # To play music on youtube
    if "Play Music".lower() in text.lower():
        text=text.replace("play music"," ")
        pywhatkit.playonyt(text)

    # To tell time
    if "the time".lower() in text.lower():
        datetime1=datetime.datetime.now().strftime("%H:%M:%S")
        speaker.Speak(f"Time is {datetime1}")

    # to send whatsapp message
    if "send whatsapp message " in text.lower():
        speaker.Speak("Sending Message")
        text.replace("send whatsapp message"," ")
        for contact in contacts:
            if contact.lower() in text.lower():
                speaker.Speak("What's the message is ")
                message=takecommand()
                pywhatkit.sendwhatmsg_instantly(contacts[contact], message, 15)
                speaker.Speak(f"Message Send to {contact}")

    # to send email
    if "send email" in text.lower():
        speaker.Speak("Sending  Email")
        text.replace("send email"," ")
        for mail in Email:
            if mail.lower() in text.lower():
                speaker.Speak("What the subject is??")
                subject=takecommand()
                speaker.Speak("What the mail is??")
                content=takecommand()
                receiver=Email[mail]
                send_email(subject,content,receiver)






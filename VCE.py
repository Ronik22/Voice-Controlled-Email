import os
from playsound import playsound
import smtplib

import speech_recognition as sr 
from gtts import gTTS


# CONSTANTS

EMAIL_ID = "5"
PASSWORD = "5"



def SpeakText(command, langinp = "en"): 
    tts = gTTS(text = command, lang = langinp)
    tts.save("~tempfile01.mp3")
    # os.system("~tempfile01.mp3")
    playsound("~tempfile01.mp3")
    print(command)
    os.remove("~tempfile01.mp3")


def speech_to_text():
    r = sr.Recognizer() 
    try: 	
        with sr.Microphone() as source2: 
            r.adjust_for_ambient_noise(source2, duration=0.2) 
            audio2 = r.listen(source2) 
                
            MyText = r.recognize_google(audio2) 
            # MyText = MyText.lower() 
            print("You said: "+MyText) 
            return MyText
                
    except sr.RequestError as e: 
        print("Could not request results; {0}".format(e)) 
        return None
            
    except sr.UnknownValueError: 
        print("unknown error occured") 
        return None


def sendMail(sendTo, msg):      # sendTo is a list []
    mail = smtplib.SMTP('smtp.gmail.com',587)    #host and port area
    mail.ehlo()  #Hostname to send for this command defaults to the FQDN of the local host.
    mail.starttls() #security connection
    mail.login(EMAIL_ID, PASSWORD) #login part
    for person in sendTo:
        mail.sendmail(EMAIL_ID, person, msg) #send part
        print ("Mail sent successfully to " + person)
    mail.close() 


def composeMail():
    SpeakText("Mention the gmail ID of the persons to whom you want to send a mail. Email IDs should be separated with the word, AND.")
    receivers = speech_to_text()
    receivers = receivers.replace("at the rate", "@")
    emails = receivers.split(" and ")
    index = 0
    for email in emails:
        emails[index] = email.replace(" ", "") 
        index += 1

    SpeakText("The mail will be send to " + (' and '.join([str(elem) for elem in emails])) + ". Confirm by saying YES or NO.")
    confirmMailList = speech_to_text()
    if confirmMailList.lower() == "yes":
        pass
    else:
        SpeakText("Operation cancelled by the user")
        return None

    SpeakText("Say your message")
    msg = speech_to_text()

    SpeakText("You said  " + msg + ". Confirm by saying YES or NO.")
    confirmMailBody = speech_to_text()
    if confirmMailBody.lower() == "yes":
        SpeakText("Message sent")
        sendMail(emails, msg)
    else:
        SpeakText("Operation cancelled by the user")
        return None


def getMailBoxStatus():
    pass

def getLatestMails():
    pass

def searchMail():
    pass

def getHelp():
    SpeakText("Choose and speak out the option number for the task you want to perform. Say 1 to send a mail. Say 2 to get your mailbox status. Say 3 to search a mail. Say 4 to get the last 3 mails. Say help to hear this message once again.")


def main():

    if EMAIL_ID != "" and PASSWORD != "":

        SpeakText("Choose and speak out the option number for the task you want to perform. Say 1 to send a mail. Say 2 to get your mailbox status. Say 3 to search a mail. Say 4 to get the last 3 mails.")
        choice = speech_to_text()

        if choice == '1' or choice.lower() == 'one':
            composeMail()
        elif choice == '2' or choice.lower() == 'too' or choice.lower() == 'two' or choice.lower() == 'to':
            getMailBoxStatus()
        elif choice == '3' or choice.lower() == 'tree' or choice.lower() == 'three':
            searchMail()
        elif choice == '4' or choice.lower() == 'four' or choice.lower() == 'for':
            getLatestMails()
        else:
            SpeakText("Wrong choice. Please say only the number")

    else:
        SpeakText("Both Email ID and Password should be present")    


if __name__ == '__main__':
    main()


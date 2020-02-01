
import os,sys
from sys import exit
def Restart(cmd):
    try:
        os.startfile(__file__)
    except:
        os.startfile(__file__[0:-3]+".exe")
    int("oof")
    

import threading, random, time,email, calendar, datetime, math,sys,webbrowser, pyttsx3, io, speech_recognition,imaplib, pickle, wmi,docx,comtypes.client,googletrans,qhue, pyowm, ast, pyttsx3.drivers, pyttsx3.drivers.sapi5, pyaudio
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from tkinter import *
from tkinter import filedialog
import smtplib
from os.path import basename
from email import parser
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.utils import COMMASPACE, formatdate
from email import encoders

windowClosed = False

#App window
class App(threading.Thread):
    def __init__(self):
        self.submitted = False
        self.window = None
        self.submit = None
        self.dots = None
        self.entryBox = None
        self.resultBox = None
        threading.Thread.__init__(self)
        self.start()
        
    def DisplayOutput(self,_text):
        self.resultBox.configure(text = _text,anchor = NW)

    def Dots(self,_text):
        self.dots.configure(text = _text)
    
    def callback(self):
        self.window.quit()

    def Submit(self):
        self.submitted = True

    def HideEntry(self):
        self.submit.grid_remove()
        self.entryBox.grid_remove()

    def ShowEntry(self):
        self.submit.grid()
        self.entryBox.grid()

    def run(self):
        self.window = Tk()
        self.window.protocol("WM_DELETE_WINDOW", self.callback)
        self.window.geometry("505x295")
        self.window.resizable(0,0)

        self.window.title("SIM Assistant")
        self.window.iconbitmap("icon0.ico")
        
        self.window.configure(background = "light grey")
        self.resultBox = Label(self.window,text="test",bg="white",fg="black",font = "none 12 bold",anchor = S,justify = LEFT,width = 50,height = 5,wraplength = 500)
        self.resultBox.grid(row=0,column=0,columnspan = 2,sticky = W)

        self.dots = Label(self.window,text=".  .  .",bg="light grey",fg="white",font = "none 20 bold",justify = LEFT,width = 26,height = 5)
        self.dots.grid(row=1,column=0,sticky=W)
        
        self.entryBox = Entry(self.window,width = 75,bg="white")
        self.entryBox.grid(row=2,column=0,sticky=W)

        self.submit = Button(self.window,text = "Ask",command = self.Submit,width = 5)
        self.submit.grid(row = 2,column = 1)

        self.HideEntry()
        Entry()

        self.window.mainloop()
        global windowClosed
        windowClosed = True

app = App()

#Assistant/misc functions

invalidCharacters = [" ","!",",","?","'","",":",";",")","(","[","]","{","}"]    
startTime = 0
stopTime = 0

try:
    assistantVoice = pickle.load(open("Voice.dat","rb"))
except Exception:
    engine = pyttsx3.init()
    assistantVoice = engine.getProperty('voices')[0]
    engine.stop()

def GetVolume():
    sessions = AudioUtilities.GetAllSessions()
    highestVol = 0
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        vol = volume.GetMasterVolume()
        if vol > highestVol:
            highestVol = vol
    return highestVol
    
previousVolume = GetVolume()

def SetVolume(vol,excludeAssistant = False):
    global previousVolume
    sessions = AudioUtilities.GetAllSessions()
    previousVolume = GetVolume()
    if type(vol) == str:
        vol = float(vol)+GetVolume()
    if vol > 1:
        vol = 1
    elif vol < 0:
        vol = 0
    for session in sessions:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        try:
            appname = session.Process.name()
        except:
            appname = ""
        if not (appname == "main.exe" or appname == "main.py"):
            volume.SetMasterVolume(vol, None)
            
muted = False
manualMode = False
didntCatch = False
executingMainThread = False
awaitingCheck = False

try:
    Email = pickle.load(open("email.dat","rb"))
except:
    Email = {}
    Email["LastEmailID"] = None
    Email["Address"] = "simassistant01@gmail.com"
    Email["Pass"] = "SimIsCool"
    Email["UserAddress"] = TypeInput("Type in your email address so I can send emails to you and vice versa.")
    pickle.dump(Email,open("email.dat","wb"))

def SendEmail(recipient,subject = "",message = "",attachment = ""):
    success = False
    attempts = 0
    while not success and attempts < 5:
        try:
            msg = MIMEMultipart()
            msg["From"] = Email["Address"]
            msg["To"] = Email["Address"]
            msg["Subject"] = subject
            msg.attach(MIMEText(message,"plain"))
            if attachment != "":
                Attachment = open(attachment,"rb")
                part = MIMEBase("application","octet-stream")
                part.set_payload((Attachment).read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition","attachement; filename ="+attachment)
                msg.attach(part)
            text = msg.as_string()
            server = smtplib.SMTP("smtp.gmail.com:587")
            server.ehlo()
            server.starttls()
            server.login(Email["Address"],Email["Pass"])
            server.sendmail(Email["Address"],recipient,text)
            server.quit()
            return True
        except Exception as e:
            attempts += 1
    return False

def Say(text,cmd = None):
    fromEmail = False
    if cmd != None and cmd != "":
        fromEmail = cmd.fromEmail
    if fromEmail:
        SendEmail(Email["UserAddress"],"Command Response",text)
    else:
        SetVolume(0.1,True)
        app.DisplayOutput(text)
        global muted
        if not muted:
            try:
                engine = pyttsx3.init()
                engine.setProperty('voice', assistantVoice)
                engine.say(text)
                engine.runAndWait()
                engine.stop()
            except Exception as e:
                app.DisplayOutput(e)
        else:
            time.sleep(len(text)*0.1+1)
        SetVolume(previousVolume)

try:
    import requests, bs4
except Exception:
    pass

Say("Loading.")

r = speech_recognition.Recognizer()

def textList(items):
    List = ""
    i = 1
    for item in items:
        List += str(i)+": "+item.capitalize()+"\n"
        i += 1
    return List

def getBody(msg):
    if msg.is_multipart():
        return getBody(msg.get_payload(0))
    else:
        return msg.get_payload(None,True)

def searchEmailPart(key,value,con):
    result, data = con.search(None,key,'"{}"'.format(value))
    return data        

# -- For remote controls
def CheckEmails():
    global awaitingCheck
    global executingmainThread
    global Email
    try:
        con = imaplib.IMAP4_SSL("imap.gmail.com")
        con.login(Email["Address"],Email["Pass"])
        con.select("INBOX")

        newestEmail = searchEmailPart("FROM",Email["UserAddress"],con)[0].split()[-1]
        if newestEmail != Email["LastEmailID"]:
            Email["LastEmailID"] = newestEmail
            result, data = con.fetch(newestEmail,"(RFC822)")
            raw = email.message_from_bytes(data[0][1])
            processed = getBody(raw).decode("utf-8")
            processed = processed.replace('\r', '')
            processed = processed.replace('\n', '')
            pickle.dump(Email,open("email.dat","wb"))
            return processed
    except:
        pass

def SpeechInput(prompt = "",cmd = None,question = False, timelimit = None):
    global stopTime
    global manualMode
    global didntCatch
    if prompt != "":
        Say(prompt,cmd)
    fromEmail = False
    if cmd != None and cmd != "":
        fromEmail = cmd.fromEmail
    if fromEmail:
        text = None
        startAsking = time.time()
        while text == None and time.time() - startAsking < 5 * 10:
            text = CheckEmails()
        if time.time() - startAsking >= 5 * 10:
            text = ""
        return text
    if not manualMode:
        while True:
            app.dots.configure(text = ".")
            text = ""
            with speech_recognition.Microphone() as source:
                try:
                    audio = r.listen(source,timelimit,10)
                except speech_recognition.WaitTimeoutError:
                    break
                stopTime = time.time()
                app.dots.configure(text = ".  .")
            try:
                text = r.recognize_google(audio)
            except  speech_recognition.UnknownValueError:
                pass
            except speech_recognition.RequestError as e:
                Say("Sorry. I'm having trouble understanding what you are saying.",cmd)
                Say("Please check your internet connection.")
                manualMode = True
                app.ShowEntry()
                text = SpeechInput(prompt)
            except Exception:
                Say("Error")
            app.dots.configure(text = ".  .  .")
            if text == "manual":
                manualMode = True
                app.ShowEntry()
                Say("Switched to manual mode.",cmd)
                text = SpeechInput(prompt)
            if text != "":
                didntCatch = False
                break   
    else:
        app.dots.configure(text = ".")
        while not app.submitted:
            pass
        app.submitted = False
        text = app.entryBox.get()
        app.entryBox.delete(0, 'end')
        app.dots.configure(text = ".  .")
        if text == "speech":
            manualMode = False
            app.HideEntry()
            Say("Switched to speech mode.",cmd)
            text = SpeechInput(prompt)
            
    if "shut down" in text:
        text = "shutdown computer"
    if "log out" in text:
        text = "logout computer"
    app.dots.configure(text = ".  .  .")
    return text    

def TypeInput(prompt):
    global manualMode
    prevMode = manualMode
    manualMode = True
    app.ShowEntry()
    result = SpeechInput(prompt)
    manualMode = prevMode
    if not manualMode:
        app.HideEntry()
    return result

def FilterAndSeperate(text,exclusion = invalidCharacters):
    result = []
    word = ""
    for char in text:
        if not char in exclusion:
            word += char
        if char == exclusion[0]:
            result.append(word)
            word = ""
    result.append(word)
    return result

def Dictionary(cmd):
    correctedWord = autocorrect.spell(cmd.cachedData[0])
    try:
        searchInput = "https://www.merriam-webster.com/dictionary/"+correctedWord
        res = requests.get(searchInput)
        res.raise_for_status()
        soup = bs4.BeautifulSoup(res.text, "html.parser")
        definitions = soup.find_all("span",attrs={"class":"dtText"})
        Say("Definition of "+correctedWord+":",cmd)
        defs = 4
        if len(definitions) < 4:
            defs = len(definitions)
        for i in range(defs):
            Say(definitions[i].text[2::].capitalize())
        
    except:
        Say("I couldn't define "+correctedWord,cmd)
    

def FilterValidAndSeperate(text,validChars):
    result = []
    word = ""
    for char in text:
        if char in validChars:
            word += char
        if char == " ":
            result.append(word)
            word = ""
    result.append(word)
    return result 

def FilterTextWords(text, validWords):
    words = FilterAndSeperate(text)
    result = []
    for word in words:
        if word in validWords:
            result.append(word)
    return result

def IsNumber(text):
    try:
        float(text)
        return True
    except Exception:
        return False
#Command Functions
#needs cmd function and can't accept other parameters

def FlipCoin(cmd):
    Say(random.choice(["Heads!","Tails!"]),cmd)

def RollDice(cmd):
    Say("You rolled a "+str(random.randint(1,6))+"!",cmd)

def Translate(cmd):
    try:
        langc = ""
        langt = ""
        for i in googletrans.LANGUAGES:
            if googletrans.LANGUAGES[i].lower() == cmd.cachedData[1]:
                langc = i
            if googletrans.LANGUAGES[i].lower() == cmd.cachedData[2]:
                langt = i      
        trans = googletrans.Translator()
        t = trans.translate(cmd.cachedData[0],langt,langc)
        Say(cmd.cachedData[0]+" in "+cmd.cachedData[2]+" is: "+t.text.capitalize(),cmd)
    except:
        Say("I couldn't translate "+cmd.cachedData[0]+" from "+cmd.cachedData[1]+" to "+cmd.cachedData[1],cmd)

def NoReply(cmd):
    Say("Hmm... I'm not sure.",cmd)

def ChangeVoice(cmd):
    global assistantVoice
    Say("Here are the voices available:",cmd)
    engine = pyttsx3.init()
    voices = engine.getProperty('voices')
    i = 0
    namedVoices = []
    i = 0
    for voice in voices:
        i += 1
        namedVoices.append(voice.id)
        engine.setProperty('voice', voice.id)
        engine.say("Option "+str(i))
        engine.runAndWait()
        engine.stop()
    try:
        chosenVoice = SpeechInput("Which option would you like to choose?",cmd)
        chosenVoice = FilterValidAndSeperate(chosenVoice,["0","1","2","3","4","5","6","7","8","9"])
        chosenVoice = int("".join(chosenVoice))
        assistantVoice = namedVoices[chosenVoice-1]
        pickle.dump(assistantVoice,open("Voice.dat","wb"))
        Say("Voice Changed to option "+str(chosenVoice),cmd)
    except Exception as e:
        print(e)
        Say("That's not an option.",cmd)
        Say("Search how to get more voices in windows if you can't find what you are looking for.",cmd)

try:
    name = pickle.load(open("Name.dat","rb"))
except Exception as e:
    name = "SIM"

username = ""

def Goodbye(cmd):
    global username
    Say("Goodbye for now, "+username+", see you soon!",cmd)
    app.callback()

def Hello(cmd):
    global username
    Say(random.choice(["Hello","Hi","Welcome","Good day","Hey"])+" "+username +", my name is " + name+".",cmd)

def ChangeUsername(cmd):
    global username
    oldusername = username
    username = SpeechInput("What do you want to change your username to?")
    if username != oldusername:
        Say(username+", I like that a little more than your last name even though "+oldusername+" was pretty good too!",cmd)
    else:
        Say("Your new username is the same as your old one. If you're having trouble saying your name as it may not be one I am familiar with you can type it in by saying 'manual' and then 'speech' when you are ready to start talking again.",cmd)
    pickle.dump(username, open("Username.dat","wb"))

def ChangeName(cmd):
    global name
    name = SpeechInput("What would you like to change my name to?",cmd)
    pickle.dump(name,open("Name.dat","wb"))
    Hello("")

def RandomNumber(cmd):
    Say("Your random number between "+ str(cmd.cachedData[0]) + " and "+ str(cmd.cachedData[1]) + " is " +str(random.randint(cmd.cachedData[0],cmd.cachedData[1]))+".",cmd)

events = {
    "christmas day": [25,12],
    "christmas": [25,12],
    "christmas eve": [24,12],
    "halloween": [31,10],
    "easter": [12,4],
    "black friday": [27,11],
    "cyber monday": [30,11],
    "boxing day": [26,12],
    "new years day": [1,1],
    "new years eve": [31,12]
    }

def WhenEvent(cmd):
    text = cmd.cachedData[0]
    text = FilterTextWords(text,["christmas","day","eve","halloween","easter","boxing","black","friday","cyber","monday","new","years"])
    try:
        event = events[" ".join(text)]
        for i in range(0,len(text)):
            text[i] = text[i].capitalize()
        text = " ".join(text)
        end = "th"
        if event[0] % 10 == 1 and event[0] != 11:
            end = "st"
        elif event[0] % 10 == 2 and event[0] != 12:
            end = "nd"
        elif event[0] % 10 == 3 and event[0] != 13:
            end = "rd"
        Say(text + " is on the "+ str(event[0]) + end + " of "+ calendar.month_name[event[1]]+".",cmd)
    except Exception as e:
        Say("I don't quite know what that event is...",cmd)

def GetDate(cmd):
    date = str(datetime.date.today())
    date = FilterAndSeperate(date,["-"])
    dow = list(calendar.day_name)[calendar.weekday(int(date[0]),int(date[1]),int(date[2]))]
    month = calendar.month_name[int(date[1])]
    end = "th"
    if date[2][len(date[2])-1] == "3" and date[2] != "11":
        end = "rd"
    elif date[2][len(date[2])-1] == "2" and date[2] != "12":
        end = "nd"
    elif date[2][len(date[2])-1] == "1" and date[2] != "13":
        end = "st"
    Say("It is "+ dow+" the "+ str(int(date[2]))+end+" of "+month + " "+ date[0],cmd)

def GetTime(cmd):
    h = int(time.strftime("%H"))
    bh = h
    h %= 12
    tod = "Am"
    if bh != h:
        tod = "Pm"
    if h == 0:
        h = 12
    m = time.strftime("%M")
    if len(m)==1:
        m = "0"+m
    Say("The time is " + str(h) + " "+ m + " "+tod+".",cmd)


try:
    Calendar = pickle.load(open("Calendar.dat","rb"))
except Exception:
    Calendar = []

def Feeling(cmd):
    Say("I am well",cmd)
    userFeeling = SpeechInput("How are you today?",cmd).lower()
    pos = ["good","well","happy","amazing","incredible","fantastic","great","lovely","really good","very good","excited"]
    neg = ["sad","mad","angry","annoyed","frustrated","unhappy","bad","terrible"]
    if userFeeling in pos:
        Say("That's great to hear.",cmd)
    elif userFeeling in neg:
        Say("Oh, that's a shame. I hope your day gets better",cmd)
    else:
        Say("Well I hope you have a great day!",cmd)

def AddCalendar(cmd):
    global Calendar
    isValid = False
    while not isValid:
        date = TypeInput("Type the date as dd/mm/yy")
        try:
            date = datetime.datetime.strptime(date, "%d/%m/%y")
            isValid = True
        except Exception as e:
            Say("The date you entered wasn't valid.")
    title = SpeechInput("What is the title you want to give this event?")
    description = SpeechInput("What description would you like to give it?")
    Calendar.append([date,title,description])
    pickle.dump(Calendar,open("Calendar.dat","wb"))
    Say("Event added to your calendar.")

try:
    HueInfo = pickle.load(open("PhilipsHue.dat","rb")) 
except Exception:
    HueInfo = {}
    pickle.dump({},open("PhilipsHue.dat","wb"))
    
def SetupHue(cmd):
    global HueInfo
    Say("Welcome to Philips Hue light setup.",cmd)
    try:
        Say("Retrieving your Hue bridge's ip...")
        searchInput = "https://www.meethue.com/api/nupnp"
        res = requests.get(searchInput)
        res.raise_for_status()
        info = bs4.BeautifulSoup(res.text, "html.parser").text
        info = ast.literal_eval(info)[0]
        HueInfo["BridgeIp"] = info["internalipaddress"]
        
        Say("Creating a username...",cmd)
        Say("Please press enter in the other window once you have pressed the button on your Hue bridge.",cmd)
        HueInfo["Username"] = qhue.create_new_username(HueInfo["BridgeIp"])
        
        Say("I would reccomend you turn on all of your Hue lights now. I will turn them off one at a time so you can give them a name to easily access them.",cmd)
        b = qhue.Bridge(HueInfo["BridgeIp"],HueInfo["Username"])
        Say("Connected to your hue bridge.",cmd)
        HueInfo["Lights"] = {}
        for light in b.lights():
            if b.lights[light].on:
                b.lights[light].state(on = False)
                lightName = SpeechInput("Find the light that I turned off. What name do you want to give this light: ",cmd).lower()
                HueInfo["Lights"][lightName] = b.lights[light]()["uniqueid"]
        Say("Saving Hue Settings...",cmd)
        pickle.dump(HueInfo,open("PhilipsHue.dat","wb")) 
        Say("Setup Complete",cmd)
    except Exception as e:
        Say("An error occured when setting up your Hue lights.",cmd)

def AddHueLight(cmd):
    global HueInfo
    if HueInfo == {}:
        Say("Hmm, It appears you don't have this command set up.",cmd)
        Say("To set up hue lights to work with me, simply say setup hue.",cmd)
        return
    try:
        b = qhue.Bridge(HueInfo["BridgeIp"],HueInfo["Username"])
        newLights = []
        for light in b.lights():
            newLight = True
            for knownLight in HueInfo["Lights"].keys():
                if HueInfo["Lights"][knownLight] == b.lights[light]()["uniqueid"]:
                    newLight = False
            if newLight:
                newLights.append(b.lights[light])
        if newLights != []:
            Say("Turn on all of your new hue lights now.",cmd)
            while not "y" in SpeechInput("Say yes when you are ready to start.",cmd):
                pass
            for light in newLights:
                light.state(on = False)
                lightName = SpeechInput("Find the light that I turned off. What name do you want to give this light: ",cmd).lower()
                HueInfo["Lights"][lightName] = light()["uniqueid"]
            Say("Saving Hue Settings...",cmd)
            pickle.dump(HueInfo,open("PhilipsHue.dat","wb"))
            Say("Complete!",cmd)
        else:
            Say("I couldn't find any new lights. If this isn't right, check you have correctly set up your Hue light",cmd)
    except:
      Say("An error occured when setting up your Hue lights.",cmd)
                    
def RenameHueLight(cmd):
    global HueInfo
    if HueInfo == {}:
        Say("Hmm, It appears you don't have this command set up.",cmd)
        Say("To set up hue lights to work with me, simply say set up hue.",cmd)
        return
    Say("Here is a list of your lights: ",cmd)
    for light in HueInfo["Lights"]:
        Say(light,cmd)
    try:
        old = SpeechInput("What is the name of the light?",cmd).lower()
        HueInfo["Lights"][old]
    except:
        Say("I couldn't find a light called "+old,cmd)
        return
    new = SpeechInput("What is the new name for the light.",cmd).lower()
    if new != old:
        HueInfo["Lights"][new] = HueInfo["Lights"][old]
        del HueInfo["Lights"][old]
        pickle.dump(HueInfo,open("PhilipsHue.dat","wb")) 
        Say("Renamed "+old+" to "+ new)
        
def HuePercent(cmd):
    global HueInfo
    if HueInfo == {}:
        Say("Hmm, It appears you don't have this command set up.",cmd)
        Say("To set up hue lights to work with me, simply say set up hue.",cmd)
        return
    try:
        lightName = cmd.cachedData[0]
        percent = cmd.cachedData[1]
        if percent < 0:
            percent = 0
        elif percent > 100:
            percent = 100
        percent = round(255 * (percent / 100))
        b = qhue.Bridge(HueInfo["BridgeIp"],HueInfo["Username"])
        lightId = ""
        for light in HueInfo["Lights"].keys():
            if lightName.lower() == light:
                lightId = HueInfo["Lights"][light]
        if lightId != "":
            for light in b.lights():
                if b.lights[light]()["uniqueid"] == HueInfo["Lights"][lightName]:
                        b.lights[light].state(bri = percent)
            Say("Turned "+lightName+" to "+str(cmd.cachedData[1])+"%",cmd)
        else:
            Say("I couldn't find a light called "+lightName,cmd)
    except:
        Say("Sorry, something went wrong",cmd)

def HueOnOff(cmd):
    global HueInfo
    if HueInfo == {}:
        Say("Hmm, It appears you don't have this command set up.",cmd)
        Say("To set up hue lights to work with me, simply say set up hue.",cmd)
        return
    try:
        lightName = cmd.cachedData[0]
        b = qhue.Bridge(HueInfo["BridgeIp"],HueInfo["Username"])
        lightId = ""
        for light in HueInfo["Lights"].keys():
            if lightName.lower() == light:
                lightId = HueInfo["Lights"][light]
        if lightId != "":
            for light in b.lights():
                if b.lights[light]()["uniqueid"] == HueInfo["Lights"][lightName]:
                    if cmd.cachedData[1] == "on":
                        b.lights[light].state(on = True)
                    else:
                        b.lights[light].state(on = False)
            Say("Turned "+ cmd.cachedData[1]+" "+lightName+".",cmd)
        else:
            Say("I couldn't find a light called "+lightName+".",cmd)
    except:
        Say("Sorry, something went wrong.",cmd)

def CheckCalendar(cmd):
    global Calendar
    try:
        Calendar = pickle.load(open("Calendar.dat","rb"))
    except Exception:
        Calendar = []
    if Calendar != []:
        found = False
        for event in Calendar:
            if event[0].date() <= datetime.datetime.now().date():
                Found = True
                Say(event[1]+"  "+event[0].strftime("%d/%m/%y"))
                Say(event[2])
                if "y" in SpeechInput("Would you like to remove this event from your calendar now?",cmd).lower():
                    Calendar.remove(event)
        pickle.dump(Calendar,open("Calendar.dat","wb"))
        if not found:
            Say("I couldn't find anything on your calendar for today.",cmd)
    else:
        Say("There isn't any events in your calendar.",cmd)

def BasicMath(cmd):
    a = cmd.cachedData[0]
    b = cmd.cachedData[1]
    r = 0
    operation = cmd.cachedData[2]
    if operation in ["add","plus","+"]:
        operation = "plus"
        r = a + b
    elif operation in ["minus","subtract","-"]:
        operation = "subtract"
        r = a - b
    elif operation in ["/","divided","over"]:
        operation = "divided by"
        r = a / b
    elif operation in ["multipied","times"]:
        operation = "multiplied by"
        r = a * b
    elif operation in ["power","^"]:
        operation = "to the power of"
        r = a ** b
    else:
        Say("I don't recognise that operation...",cmd)
        return
    try:
        if r % 1 == 0:
            r = int(r)
        else:
            round(r)
    except Exception:
        pass
    Say("The result of "+str(a)+" "+operation+" "+str(b)+ " is "+ str(r)+".",cmd)
    

def CorrectPath(path):
    newPath = ""
    for i in range(0,len(path)):
        if path[i] == "\\":
            newPath += "/"
        else:
            newPath += path[i]
    return newPath

def Joke(cmd):
    from Joke import jokeBook
    joke = random.choice(list(jokeBook.keys()))
    Say(joke,cmd)
    Say(jokeBook[joke],cmd)


def GetDirectory():
    direct = ""
    while True:
        direct = filedialog.askopenfilename(initialdir  ="/", filetypes = [("All files","*.*")])
        if "y" in SpeechInput("Are you sure you want to set this as the directory?").lower():
            break
    
    return direct

try:
    appDirs = pickle.load(open("AppInfo.dat","rb"))
except Exception:
    appDirs = {}

def ChangeApp(cmd):
    global appDirs
    appName = SpeechInput("What is the name of the app you would like to create or change?",cmd).lower()
    appName = "".join(FilterAndSeperate(appName,[" "])).lower()
    if appName in list(appDirs.keys()):
        if "y" in SpeechInput("An app called "+appName+" already exists, are you sure you want to change it?",cmd).lower():
            appDirs[appName] = GetDirectory()
            pickle.dump(appDirs,open("AppInfo.dat","wb"))
            Say("Changed app called"+appName+" successfully.",cmd)
    else:
        if "y" in SpeechInput("Are you sure you want to add an app called "+appName+".",cmd).lower():
            appDirs[appName] = GetDirectory()
            pickle.dump(appDirs,open("AppInfo.dat","wb"))
            Say("Added app called"+appName+" successfully.",cmd)

c = wmi.WMI()
try:
    ignoredDirs = pickle.load(open("IgnoredDirectories.dat","rb"))
except Exception:
    ignoredDirs = []
    for p in c.Win32_Process():
        ignoredDirs.append(p.ExecutablePath)
    pickle.dump(ignoredDirs,open("IgnoredDirectories.dat","wb")) 
   
def CheckOpenApps():
    global appDirs
    global ignoredDirs
    global awaitingCheck
    global executingmainThread
    c = wmi.WMI()
    Apps = c.Win32_Process()
    for p in Apps:
        if not p.ExecutablePath in list(appDirs.values()) and not p.ExecutablePath in ignoredDirs:
            awaitingCheck = True
            while executingMainThread:
                pass
            if "y" in SpeechInput("I noticed an app I haven't seen before called "+ p.Name+". Would you like to create a shortcut to it?").lower():
                appName = SpeechInput("What would you like to call it?")
                appName = "".join(FilterAndSeperate(appName,[" "])).lower()
                if appName in list(appDirs.keys()):
                    if "y" in SpeechInput("An app called "+appName+" already exists, are you sure you want to change it?").lower():
                        appDirs[appName] = p.ExecutablePath
                        pickle.dump(appDirs,open("AppInfo.dat","wb"))
                else:
                    if "y" in SpeechInput("Are you sure you want to add an app called "+appName+".").lower():
                        appDirs[appName] = p.ExecutablePath
                        pickle.dump(appDirs,open("AppInfo.dat","wb"))
            else:
                ignoredDirs.append(p.ExecutablePath)
                pickle.dump(ignoredDirs,open("IgnoredDirectories.dat","wb"))
    awaitingCheck = False


def OpenApp(cmd):
    global appDirs
    appName = cmd.cachedData[0].replace(" ","")
    selectedApp = ""
    for app in appDirs:
        if app == appName:
            selectedApp = appDirs[app]
    if selectedApp != "":
        Say("Opening "+ appName,cmd)
        if not os.path.isfile(selectedApp):
            Say("The directory to "+ selectedApp + "is incorrect or has been changed.",cmd)
            Say("You can now change it to the correct directory.",cmd)
            appDirectory = GetDirectory()
            selectedApp = appDirectory
            appDirs[appName] = appDirectory
            pickle.dump(appDirs,open("AppInfo.dat","wb"))
        os.startfile(selectedApp)
    else:
        Say("I couldn't find an app called " + appName+".",cmd)
        if "y" in SpeechInput("Would you like to create an app called "+appName+"?: ",cmd).lower():
            appDirectory = GetDirectory()
            appDirs[appName] = appDirectory
            pickle.dump(appDirs,open("AppInfo.dat","wb"))
            Say("Successfully added "+ appName+".",cmd)
            os.startfile(selectedApp)
     
def StartTimer(cmd):
    global startTime
    startTime = time.time()
    Say("Timer started!",cmd)

def StopTimer(cmd):
    global startTime
    global stopTime
    if startTime != 0:
        seconds = stopTime - startTime
        startTime = 0
        hours = int(math.floor(seconds / (3600)))
        seconds -= hours * 3600
        minutes = int(math.floor(seconds / (60)))
        seconds -= minutes * 60
        milliseconds = round((seconds - int(math.floor(seconds))) * 1000)
        seconds = int(math.floor(seconds))
        Say("Timer Stopped:",cmd)
        Say("Hours: "+str(hours)+", Minutes: "+ str(minutes)+", Seconds: "+ str(seconds)+", Milliseconds: "+str(milliseconds),cmd)
    else:
        if "y" in SpeechInput("You never started a timer. Would you like to start a timer now?: ",cmd).lower():
            StartTimer(cmd)

try:
    userLocation = pickle.load(open("location.dat","rb"))
except Exception:
    while True:
        userLocation = SpeechInput("For tasks that require infomation local to where you live, what town do you live in?",cmd)
        if "y" in SpeechInput("Are you sure you want to set "+userLocation+" as your location?",cmd).lower():
            break 
                        
    pickle.dump(userLocation,open("location.dat","wb"))
def Weather(cmd):
    try:
        owm = pyowm.OWM("eba6ddc132a79ff6df803747ddf34df7")
        location = owm.weather_at_place(userLocation)
        weather = location.get_weather()

        temp = weather.get_temperature("celsius")["temp"]
        status =weather.get_status().lower() #clouds/clear/mist/rain/snow
        
        text = "Today in "+userLocation+", It is "
        suggestion = ""
        if temp <= 0:
            suggestion = "Its going to be freezing today so wrap up really warm"
        elif (temp < 15 and temp > 0) or status == "snow":
            suggestion= "Wrap up warm. It should be cold today"
        elif temp > 22 and temp <= 30:
            suggestion= "It should be nice and hot today"
        elif temp > 30:
            suggestion= "Its very hot today. Remember to drink lots of water to stay hydrated"
        else:
            suggestion = "The temperature should be quite mild."
        if status == "clouds":
            text += "cloudy"
        elif status == "mist":
            suggestion += " and be careful if you're driving."
            text += "misty"
        elif status == "rain" or status == "drizzle":
            suggestion += " and don't forget a coat and umbrella."
            text += "raining"
        elif status == "snow":
            suggestion += " and don't forget to wear a coat. Have fun in the snow!"
            text += "snowing"
        else:
            text += status
        text += " at a temperature of "+str(round(temp))+" degrees celsius. " + suggestion
        Say(text,cmd)
    except:
        Say("I am not sure what the weather is right now.",cmd)
    
def Search(cmd):
    try:
        searchInput = cmd.cachedData[0]
        searchInput = "https://google.com/search?q="+searchInput
        res = requests.get(searchInput)
        res.raise_for_status()
        soup = bs4.BeautifulSoup(res.text, "html.parser")
        linkElements = soup.select('div#main > div > div > div > a')
        if len(linkElements) == 0:
            Say("Couldn't find any results...",cmd)
            return
        link = linkElements[0].get("href")
        i = 0
        while link[0:4] != "/url" or link[14:20] == "google":
            i += 1
            link = linkElements[i].get("href")
        Say("Searching google for " + cmd.cachedData[0]+".",cmd)
        webbrowser.open("http://google.com"+link)
    except Exception as e:
        Say("Something went wrong while searching for "+ cmd.cachedData[0]+".",cmd)

try:
    shoppingList = pickle.load(open("shoppingList.dat","rb"))
except:
    shoppingList = []
    
def AddShoppingList(cmd):
    global shoppingList
    shoppingList.append(cmd.cachedData[0])
    pickle.dump(shoppingList,open("shoppingList.dat","wb"))
    Say("Added "+cmd.cachedData[0]+" to your shopping list.",cmd)

    
def RemoveShoppingList(cmd):
    global shoppingList
    try:
        shoppingList.remove(cmd.cachedData[0])
        pickle.dump(shoppingList,open("shoppingList.dat","wb"))
        Say("Removed "+cmd.cachedData[0]+" to your shopping list.",cmd)
    except:
        pass

def ClearShoppingList(cmd):
    global shoppingList
    try:
        shoppingList = []
        pickle.dump(shoppingList,open("shoppingList.dat","wb"))
        Say("Cleared your shopping list.",cmd)
    except:
        pass

def GoShopping(cmd):
    if shoppingList != []:
        try:
            for item in shoppingList:
                searchInput = "https://google.com/search?q="+item+" amazon"
                res = requests.get(searchInput)
                res.raise_for_status()
                soup = bs4.BeautifulSoup(res.text, "html.parser")
                linkElements = soup.select('div#main > div > div > div > a')
                if len(linkElements) == 0:
                    continue
                link = linkElements[0].get("href")
                i = 0
                while link[0:4] != "/url" or link[14:20] == "google":
                    i += 1
                    link = linkElements[i].get("href")
                webbrowser.open("http://google.com"+link)
            if "y" in SpeechInput("Would you like to clear your shopping list?",cmd).lower():
                ClearShoppingList("")
        except Exception as e:
            Say("You can't go shopping right now.",cmd)
    else:
        Say("You don't have anything on your shopping list.",cmd)

def CMDEmail(cmd):
    recipient = TypeInput("Type the email address of the email recipient:",cmd)
    subject = SpeechInput("Enter a subject for the email:",cmd)
    message = SpeechInput("Enter in a message:",cmd)
    Say("Add an attachment:",cmd)
    attachment = GetDirectory()
    
    success = SendEmail(recipient,subject,message,attachment)
    if success:
        Say("Email Sent")
    else:
        Say("Email Failed To Send")
        
def PrintShoppingList(cmd):
    try:
        List = textList(shoppingList)

        SendEmail(Email["UserAddress"],"Shopping List",List)
    except:
        Say("Failed to send your shopping list.",cmd)

def NumericalFilter(text,validCharacters = ["0","1","2","3","4","5","6","7","8","9","-",".","+"]):
    result = ""
    for char in text:
        if char in validCharacters:
            result += char
    return result

def Mute(cmd):
    global muted
    muted = True
    Say("Muted.",cmd)

def Unmute(cmd):
    global muted
    muted = False
    Say("Unmuted.",cmd)

def CMDSetVolume(cmd):
    volume = cmd.cachedData[0]
    volume = cmd.cachedData[0] / 100
    SetVolume(volume)
    Say("Set volume to "+str(int(volume * 100))+"%",cmd)

def VolumeUpDown(cmd):
    if cmd.cachedData[0] == "up":
        SetVolume("0.1")
    elif cmd.cachedData[0] == "down":
        SetVolume("-0.1")
    else:
        return
    Say("Turned volume "+cmd.cachedData[0]+".",cmd)
        

def Marsh(cmd):
    Say("Mr marsh is life.",cmd)

lastCMD = None
lastFilterInput = []
CMDs = {}
CMDGroups = {}

def Help(cmd):
    while True:
        Say("Here are the available topics:",cmd)
        if not cmd.fromEmail:
            Say("- Commands",cmd)
            Say("- How I Work",cmd)
            Say("- Talking",cmd)
        else:
            Say(textList(["Commands", "How I work", "Talking"]),cmd)
        Topic = SpeechInput("What Topic Would you like me to cover?",cmd).lower()

        if Topic == "commands":
            Say("Here are the available command topics:",cmd)
            if not cmd.fromEmail:
                for group in CMDGroups:
                    if group != "***":
                        Say("- "+group.capitalize())
            else:
                Say(textList(CMDGroups))
            commandHelp = True
            group = SpeechInput("What command topic would you like help with:",cmd).lower()
            if group in CMDGroups.keys() and group != "***":
                Say("Here are the available commands:",cmd)
                if not cmd.fromEmail:
                    for command in CMDGroups[group]:
                        Say("- "+command.capitalize())
                else:
                    Say(textList(CMDGroups[group]))
                while commandHelp:
                    query = SpeechInput("What would you like help with? (say done to leave help)",cmd).lower()
                    if query == "done":
                        commandHelp = False
                    else:
                        try:
                            Say(CMDGroups[group][query].help)
                            Say("Required Keywords:",cmd)
                            if not cmd.fromEmail:
                                app.DisplayOutput(str(CMDGroups[group][query].reqWords))
                                time.sleep(4)
                            else:
                                Say(str(CMDGroups[group][query].reqWords))
                        except Exception as e:
                            print(e)
                            Say("I couldn't help you with the query, "+query,cmd)
        elif Topic == "how i work" or Topic == "how you work":
            Say("I work by analysing each word you tell me.",cmd)
            Say("I have a certain set of commands of which some have required keywords that I use to figure out what you want me to do.",cmd)
            Say("Some keywords help tell me if I need to be looking for infomation.",cmd)
            Say("such as the open keyword that tells me the name of the app will come after it.",cmd)
            Say("This can help give you more flexibility in how you say certain commands as I only look for keywords.",cmd)
            Say("Although I can get a bit confused if you don't use the required keywords.",cmd)
            Say("I work out what you are saying by using Google's speech to text API which requires an internet connection.",cmd)
            Say("I can't work out what you say to me without a connection and so you will be limited to typing in commands.",cmd)
        elif Topic == "talking":
            Say("When talking to me you may notice several dots appearing on screen.",cmd)
            Say("1 dot means I am listening for your command.",cmd)
            Say("2 dots means I am working out what you said to me.",cmd)
            Say("3 dots means I am ready to carry out your given command.",cmd)
            Say("I can't hear you while I am working out or carrying out a command.",cmd)
            Say("So make sure you only talk to me while there is 1 dot on screen.",cmd)
            Say("Also, in some cases you might want to type what you want me to do.",cmd)
            Say("If so you can switch between the two modes by saying 'manual' and 'speech'",cmd)
        else:
            Say(Topic+" is not a valid topic.",cmd)
        if "n" in SpeechInput("Would you like to continue?",cmd).lower():
            break
    Say("Bye for now! Remember, I am always here to help you!",cmd)
                
#Command Handling

class cmdWord:
    def __init__(self,startRecording = False,stopRecording = False,addCMDWord = False,order = 0):
        self.order = order
        self.includeWord = addCMDWord
        self.start = startRecording
        self.stop = stopRecording

class cmdType:
    def __init__(self,numerical = False,multiWord = True):
        self.numerical = numerical
        self.multiWord = multiWord

def AddCachedData(Data,CachedData,dataFormat):
    for i in range(0, len(dataFormat)):
        if CachedData[i] == "":
            if dataFormat[i].numerical == IsNumber(Data) and (dataFormat[i].multiWord == (len(FilterAndSeperate(Data)) > 1) or dataFormat[i].multiWord):
                if dataFormat[i].numerical:
                    try:
                        CachedData[i] = int(Data)
                    except Exception:
                        CachedData[i] = float(Data)
                else:
                    CachedData[i] = Data
                return CachedData
    return CachedData

def AmountEmptyData(cachedData):
    empty = 0
    for data in cachedData:
        if data == "":
            empty += 1
    return empty
                
def FindEmptyData(cmd,cachedData):
    for i in range(0,len(cmd.dataFormat)):
        if cachedData[i] == "":
            return cmd.dataFormat[i]
    return True

class CMD:
    def __init__(self,group,name,weight,requiredCmdWords = [[""]],function = NoReply,helpDescription = "",dataFormat = [cmdType()],cmdWords = {},getDataOnStart = False):
        if group in CMDGroups.keys():
            CMDGroups[group][name] = self
        else:
            CMDGroups[group] = {}
            CMDGroups[group][name] = self
        self.fromEmail = False
        self.weight = weight
        self.getDataOnStart = getDataOnStart
        self.cmdWords = cmdWords
        self.dataFormat = dataFormat
        self.help = helpDescription
        self.name = name
        self.function = function
        self.reqWords = requiredCmdWords
        self.cachedData = []
        for i in range(0,len(dataFormat)):
            self.cachedData.append("")
        CMDs[name] = self
    def __repr__(self):
        return self.name
    def relevantCommand(self,userInput):
        found = []
        for word in userInput:
            for i in range(0,len(self.reqWords)):
                if word in self.reqWords[i] and not self.reqWords[i] in found:
                    found.append(self.reqWords[i])
        if len(found) == len(self.reqWords):
            return True
        return False
    def run(self,words):
        iterations = 0
        while FindEmptyData(self,self.cachedData) != True and iterations < 64:
            iterations += 1
            GetData = self.getDataOnStart
            criteria = FindEmptyData(self,self.cachedData)
            if criteria == True:
                break
            Data = ""
            cmdWordIndex = 0
            for word in words:
                IncrementedIndex = False
                if word in self.cmdWords.keys():
                    if (self.cmdWords[word].stop and self.cmdWords[word].order == cmdWordIndex):
                        IncrementedIndex = True
                        cmdWordIndex += 1
                        GetData = False
                        self.cachedData = AddCachedData(Data,self.cachedData,self.dataFormat)
                        Data = ""
                        criteria = FindEmptyData(self,self.cachedData)
                        if criteria == True:
                            break
                    if self.cmdWords[word].includeWord and (self.cmdWords[word].order == cmdWordIndex or IncrementedIndex):
                        IncrementedIndex = True
                        if not IncrementedIndex:
                            cmdWordIndex += 1
                        cachedData = AddCachedData(word,self.cachedData,self.dataFormat)
                        criteria = FindEmptyData(self,self.cachedData)
                        if criteria == True:
                            break
                if not criteria.multiWord and Data != "":
                    GetData = False
                if GetData:
                    if Data != "" and criteria.multiWord:
                        Data += " "
                    if criteria.numerical:
                        Data += NumericalFilter(word)
                    else:
                        Data += word
                if word in self.cmdWords.keys():
                    if self.cmdWords[word].start and (self.cmdWords[word].order == cmdWordIndex or IncrementedIndex):
                        if not IncrementedIndex:
                            cmdWordIndex += 1
                        GetData = True
            self.cachedData = AddCachedData(Data,self.cachedData,self.dataFormat)
        if FindEmptyData(self,self.cachedData) == True:
            self.function(self)
        self.cachedData = []
        for i in range(0,len(self.dataFormat)):
            self.cachedData.append("")

def Repeat(cmd):
    if lastCMD != None:
        lastCMD.run(lastFilterInput)
    else:
        NoReply(cmd)

def BestCMD(Words):
    best = str(CMDs[list(CMDs.keys())[0]])
    for cmd_ in CMDs:
        if CMDs[cmd_].relevantCommand(Words) and ((len(CMDs[cmd_].reqWords) > len(CMDs[best].reqWords) and CMDs[cmd_].weight >= CMDs[best].weight) or CMDs[cmd_].weight > CMDs[best].weight):
            best = cmd_
    return CMDs[best]

def Power(cmd):
    if cmd.cachedData[0] == "shutdown":
        if "y" in SpeechInput("Are you sure you want to shutdown your computer?").lower():
            Say("Shutting Down",cmd)
            os.system("shutdown /s")
            return 
    elif cmd.cachedData[0] == "restart":
        if "y" in SpeechInput("Are you sure you want to restart your computer?",cmd).lower():
            Say("Restarting",cmd)
            os.system("shutdown /r")
            return 
    elif cmd.cachedData[0] == "logout":
        if "y" in SpeechInput("Are you sure you want to log out of your computer?",cmd).lower():
            Say("Logging Out",cmd)
            os.system("shutdown /l")
            return
    Say(cmd.cachedData[0].capitalize()+" cancelled.",cmd)

#CommandWords


#Commands
CMD("***","no reply",0,[],NoReply,"This just tells you that I didn't quite understand what you wanted me to do.",[])
CMD("***","help",0,["help"],Help,"Provides help.",[])
CMD("basic","repeat",0,[["repeat"]],Repeat,"Repeats the last command.",[])
CMD("basic","hello",0,[["hello","hi","hey"]],Hello,"Simply say hello.",[])
CMD("basic","goodbye",0,[["goodbye","bye"]],Goodbye,"Simply say goodbye to close me.",[])
CMD("personalisation","change my name",1,[["change"],["your"],["name","username"]],ChangeName,"Change my name to something that feels more personal.",[])
CMD("personalisation","change your name",1,[["change"],["my"],["name","username"]],ChangeUsername,"Change your username to something that feels more personal.",[])
CMD("personalisation","change voice",1,[["change"],["voice"]],ChangeVoice,"Change my voice to one of several option that are demoed to you beforehand.",[])
CMD("control","open app",1,[["open"]],OpenApp,"Open an app by telling me its name. If I don't already know where that app is I'll ask you for its directory so I can create a shortcut.",[cmdType(False,True)],{"open": cmdWord(True,False)})
CMD("control","change app",1,[["change","add"],["app"]],ChangeApp,"If you made a mistake when giving me an app's directory or want to add an new app's directory.",[])
CMD("internet","google search",3,[["search", "google"]],Search,"Search google using the 'search' or 'google' keyword and I will provide you with the top search result.",[cmdType()],{"search":cmdWord(True),"google":cmdWord(True)})
CMD("basic","joke",0,[["joke"]],Joke,"I'll tell you a joke.",[])
CMD("basic","time",0,[["time"]],GetTime,"I'll tell you the time.",[])
CMD("basic","date",0,[["date","day","today"]],GetDate,"I'll tell you the date.",[])
CMD("numbers","random number",1,[["number"],["and","to"]],RandomNumber,"I'll give you a random number between a and b which could be set out similar to the following: give me a random number between a and b.",[cmdType(True,False),cmdType(True,False)],{"number":cmdWord(True,False,False,0),"to":cmdWord(True,True,False,1),"and":cmdWord(True,True,False,1)})
CMD("numbers","simple math",1,[["add","plus","subtract","minus","times","power","^","multiplied","divided","/","over","+","-"]],BasicMath,"I can do division, multiplication, subtraction and addition of two numbers'",[cmdType(True,False),cmdType(True,False),cmdType(False,False)],{"/":cmdWord(True,True,True),"add":cmdWord(True,True,True),"plus":cmdWord(True,True,True),"subtract":cmdWord(True,True,True),"minus":cmdWord(True,True,True),"times":cmdWord(True,True,True),"power":cmdWord(True,True,True),"^":cmdWord(True,True,True),"multiplied":cmdWord(True,True,True),"divided":cmdWord(True,True,True),"over":cmdWord(True,True,True),"+":cmdWord(True,True,True),"-":cmdWord(True,True,True)},True)
CMD("numbers","start timer",1,[["start","begin"],["timer"]],StartTimer,"I'll start a timer.",[])
CMD("numbers","stop timer",1,[["stop","end","finish"],["timer"]],StopTimer,"I'll stop a timer and tell you how long it lasted.",[])
CMD("numbers","events",1,[["when"],["is"]],WhenEvent,"I can tell you when special events like halloween are.")
CMD("personalisation","add calendar",1,[["add"],["calendar"]],AddCalendar,"Set events on certain dates to remind you to do tasks and stay organised.",[])
CMD("personalisation","check calendar",1,[["check","show"],["calendar"]],CheckCalendar,"Check tasks you have to do today on your calendar to remind you to do tasks and stay organised.",[])
CMD("***","???",69,[["mr"],["marsh"],["is"],["love"]],Marsh,"???",[])
CMD("control","mute",0,[["mute"]],Mute,"Mute my voice but text will still appear on screen.",[])
CMD("control","unmute",0,[["unmute"]],Unmute,"Unmute my voice when you want.",[])
CMD("shopping","add shopping list",1,[["add"],["to"],["shopping"],["list"]],AddShoppingList,"Add something to your shopping list.",[cmdType(False,True)],{"add":cmdWord(True,True,False,0),"to":cmdWord(False,True,False,1)})
CMD("shopping","remove shopping list",1,[["remove"],["to","from"],["shopping"],["list"]],RemoveShoppingList,"Remove an item from your shopping list.",[cmdType(False,False),cmdType(False,True)],{"remove":cmdWord(True,False,False,0),"to":cmdWord(False,True,False,1),"from":cmdWord(False,True,False,1)})
CMD("shopping","clear shopping list",1,[["clear"],["shopping"],["list"]],ClearShoppingList,"Removes all items from shopping your list.",[])
CMD("shopping","go shopping",1,[["go"],["shopping"]],GoShopping,"Searches each item on your shopping list on amazon.",[])
CMD("shopping","email shopping list",1,[["email","send","give","get","print"],["shopping"],["list"]],PrintShoppingList,"Emails your shopping list to the email that you provided me with.",[])
CMD("internet","email",0,[["send","write"],["email"]],CMDEmail,"Sends an email through Gmail.",[])
CMD("internet","translate",2,[["translate"],["from"],["to"]],Translate,"Translate some words from one language to another",[cmdType(False,True),cmdType(False,False),cmdType(False,False)],{"translate":cmdWord(True,False,False,0),"from":cmdWord(True,True,False,1),"to":cmdWord(True,True,False,2)})
CMD("basic","weather",1,[["weather"]],Weather,"I'll tell you waht the weather is like today and suggest what to wear.",[])
CMD("control","power",0,[["shutdown","logout","restart"],["computer"]],Power,"Allows you to restart/shutdown/logout simply by asking.",[cmdType(False,False)],{"shutdown":cmdWord(False,False,True),"logout":cmdWord(False,False,True),"restart":cmdWord(False,False,True)})
CMD("lighting","hue setup",0,[["set"],["up"],["hue","lights","light"]],SetupHue,"Allows you to use your Philips Hue lights with me.",[])
CMD("lighting","hue turn on off",1,[["turn"],["on","off"]],HueOnOff,"Turn a Philips Hue light on or off.",[cmdType(False,True),cmdType(False,False)],{"turn":cmdWord(True,False,False,0),"on":cmdWord(False,True,True,1),"off":cmdWord(False,True,True,1)})
CMD("lighting","rename hue light",1,[["rename"],["hue","light","lights"]],RenameHueLight,"Rename a light to something that better suits it. This name is used to say what light should be turned on/off etc.",[])
CMD("lighting","hue light brightness",2,[["turn"],["to"]],HuePercent,"Set the brightness of a given light as a given percent.",[cmdType(False,True),cmdType(True,False)],{"turn":cmdWord(True,False,False,0),"on":cmdWord(False,True,False,1),"to":cmdWord(True,True,False,2),"up":cmdWord(False,True,False,1),"down":cmdWord(False,True,False,1)})
CMD("basic","flip coin",0,[["flip"],["coin"]],FlipCoin,"I'll flip a coin for you incase you want to settle an argument.",[])
CMD("basic","roll dice",0,[["roll"],["die","dice"]],RollDice,"I can roll a die if you don't have one.",[])
CMD("lighting","add hue light",1,[["add"],["light","lights","hue"]],AddHueLight,"Add a new hue light to your setup",[])
CMD("control","restart",0,[["restart","assistant"]],Restart,"Restart me.",[])
CMD("basic","how are you",1,[["how"],["are"],["you"]],Feeling,"Ask me how I feel and I'll ask you back.",[])
CMD("control","set volume",1,[["volume"]],CMDSetVolume,"Set a volume for apps on your computer to a certain percentage.",[cmdType(True,False)],{"volume":cmdWord(True)})
CMD("control","volume up down",1,[["volume"],["up","down"]],VolumeUpDown,"Turn volume up or down by 10%",[cmdType(False,False)],{"up":cmdWord(False,False,True),"down":cmdWord(False,False,True)})
CMD("internet","define",1,[["define"]],Dictionary,"I'll define a word with the help of dictionary.com",[cmdType(False,False)],{"define":cmdWord(True)})
    
#42 commands

appCheckFrequency = 1 * 60

class Checker(threading.Thread):
    def run(self):
        CheckOpenApps()        
        global userInput
        lastCheck = time.time()
        while True:
            emailInput = CheckEmails()
            if emailInput != None:
                emailInput = emailInput.lower()
                filterInput = FilterAndSeperate(emailInput)
                UserCMD = BestCMD(filterInput)
                UserCMD.fromEmail = True
                UserCMD.run(filterInput)
                lastFilterInput = filterInput
                if UserCMD.function != Repeat:
                    lastCMD = UserCMD
            if time.time() - lastCheck >= appCheckFrequency:
                CheckOpenApps()
                lastCheck = time.time()

Checker().start()

try:
    username = pickle.load(open("Username.dat","rb"))
    Hello("")
except Exception:
    Say("Hello I don't believe we've met before!")
    username = SpeechInput("What's your name?")
    Say("Hello, "+username+"! That's a nice name. my name is "+ name+ " but you can change it if you want.")
    pickle.dump(username, open("Username.dat","wb"))
    Say("I recommend you start off in help to see what you can do.")
    Help("")

try:
    while True:
        print(awaitingCheck)
        while awaitingCheck:
            time.sleep(0.2)
        executingMainThread = True
        app.dots.configure(text = ".")
        if not manualMode or (manualMode and app.submitted):
            userInput = SpeechInput("",None,False,10)
            if userInput != "" and userInput != "speech" and userInput != "manual":
                userInput = userInput.lower()
                filterInput = FilterAndSeperate(userInput)
                UserCMD = BestCMD(filterInput)
                UserCMD.run(filterInput)
                lastFilterInput = filterInput
                if UserCMD.function != Repeat:
                    lastCMD = UserCMD
        if windowClosed:
            exit()
        executingMainThread = False
                                
except BaseException as e:
    print(e)
    Say("An error occured.")
    app.dots.configure(text = "ERROR: "+str(e))
    Say("Automatically restarting in 5 seconds.")
    time.sleep(5)
    Restart("")

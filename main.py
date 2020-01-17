try:
    import os,sys
    from sys import exit

    def Restart(cmd):
        try:
            os.startfile(__file__)
        except:
            os.startfile(__file__[0:-3]+".exe")
        exit()
    import threading, random, time, calendar, datetime, math,sys,webbrowser, pyttsx3, io, speech_recognition, pickle, wmi,docx,comtypes.client,googletrans,qhue, pyowm, ast, pyttsx3.drivers, pyttsx3.drivers.sapi5
    from tkinter import *
    from tkinter import filedialog
    import smtplib
    from os.path import basename
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
            exit()

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
            self.window.iconbitmap("icon.ico")
            
            self.window.configure(background = "light grey")
            self.resultBox = Label(self.window,text="test",bg="white",fg="black",font = "none 12 bold",anchor = S,justify = LEFT,width = 50,height = 5,wraplength = 500)
            self.resultBox.grid(row=0,column=0,columnspan = 2,sticky = W)

            self.dots = Label(self.window,text=".  .  .",bg="light grey",fg="white",font = "none 20 bold",justify = LEFT,width = 26,height = 5)
            self.dots.grid(row=1,column=0,sticky=W)
            
            self.entryBox = Entry(self.window,width = 75,bg="white")
            self.entryBox.grid(row=2,column=0,sticky=W)

            self.submit = Button(self.window,text = "Ask",command = self.Submit,width = 5)
            self.submit.grid(row = 2,column = 1)

            self.ShowEntry()

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
        
    muted = False

    def Say(text):
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

    try:
        import requests, bs4
    except Exception:
        Say("Error: The requests module cannot work properly. This might cause web based commands to not work correctly")

    Say("Loading")

    r = speech_recognition.Recognizer()

    manualMode = True
    didntCatch = False
    def SpeechInput(prompt = "",question = False, timelimit = None):
        global stopTime
        global manualMode
        global didntCatch
        if prompt != "":
            Say(prompt)
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
                    Say("Sorry. I'm having trouble understanding what you are saying.")
                    Say("Please check your internet connection.")
                    manualMode = True
                    showEntry()
                    text = SpeechInput(prompt)
                except Exception:
                    say("Error")
                app.dots.configure(text = ".  .  .")
                if text == "manual":
                    manualMode = True
                    app.ShowEntry()
                    Say("Switched to manual mode")
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
                Say("Switched to speech mode")
                text = SpeechInput(prompt)
                
        if "shut down" in text:
            text = "shutdown computer"
        if "log out" in text:
            text = "logout computer"
        app.dots.configure(text = ".  .  .")
        return text    

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
        Say(random.choice(["Heads!","Tails!"]))

    def RollDice(cmd):
        Say("You rolled a "+str(random.randint(1,6))+"!")

    def Translate(cmd):
        langc = ""
        langt = ""
        for i in googletrans.LANGUAGES:
            if googletrans.LANGUAGES[i].lower() == cmd.cachedData[1]:
                langc = i
            if googletrans.LANGUAGES[i].lower() == cmd.cachedData[2]:
                langt = i      
        trans = googletrans.Translator()
        t = trans.translate(cmd.cachedData[0],langt,langc)
        Say(cmd.cachedData[0]+" in "+cmd.cachedData[2]+" is")
        app.DisplayOutput(t.text)

    def NoReply(cmd):
        Say("Hmm... I'm not sure.")

    def ChangeVoice(cmd):
        global assistantVoice
        Say("Here are the voices available:")
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
            chosenVoice = SpeechInput("Which option would you like to choose?")
            chosenVoice = FilterValidAndSeperate(chosenVoice,["0","1","2","3","4","5","6","7","8","9"])
            chosenVoice = int("".join(chosenVoice))
            assistantVoice = namedVoices[chosenVoice-1]
            pickle.dump(assistantVoice,open("Voice.dat","wb"))
            Say("Voice Changed to option "+str(chosenVoice))
        except Exception:
            Say("That's not an option")
            Say("Search how to get more voices in windows if you can't find what you are looking for")

    try:
        name = pickle.load(open("Name.dat","rb"))
    except Exception as e:
        name = "SIM"

    username = ""

    def Goodbye(cmd):
        global username
        Say("Goodbye for now, "+username+", see you soon!")
        exit()

    def Hello(cmd):
        global username
        Say(random.choice(["Hello","Hi","Welcome","Good day","Hey"])+" "+username +", my name is " + name)

    def ChangeUsername(cmd):
        global username
        oldusername = username
        username = SpeechInput("What do you want to change your username to?")
        if username != oldusername:
            Say(username+", I like that a little more than your last name even though "+oldusername+" was pretty good too!")
        else:
            Say("Your new username is the same as your old one. If you're having trouble saying your name as it may not be one I am familiar with you can type it in by saying 'manual' and then 'speech' when you are ready to start talking again.")
        pickle.dump(username, open("Username.dat","wb"))

    def ChangeName(cmd):
        global name
        name = SpeechInput("What would you like to change my name to?")
        if name.lower() == "manual":
            name = input(": ")
        pickle.dump(name,open("Name.dat","wb"))
        Hello("")

    def RandomNumber(cmd):
        Say("Your random number between "+ str(cmd.cachedData[0]) + " and "+ str(cmd.cachedData[1]) + " is " +str(random.randint(cmd.cachedData[0],cmd.cachedData[1])))

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
            Say(text + " is on the "+ str(event[0]) + end + " of "+ calendar.month_name[event[1]] )
        except Exception as e:
            Say("I don't quite know what that event is...")

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
        Say("It is "+ dow+" the "+ str(int(date[2]))+end+" of "+month + " "+ date[0])

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
        Say("The time is " + str(h) + " "+ m + " "+tod)


    try:
        Calendar = pickle.load(open("Calendar.dat","rb"))
    except Exception:
        Calendar = []
        
    def AddCalendar(cmd):
        global Calendar
        isValid = False
        while not isValid:
            Say("Type the date as dd/mm/yy")
            date = input(": ")
            try:
                date = datetime.datetime.strptime(date, "%d/%m/%y")
                isValid = True
            except Exception as e:
                Say("The date you entered wasn't valid.")
        title = SpeechInput("What is the title you want to give this event?")
        description = SpeechInput("What description would you like to give it?")
        Calendar.append([date,title,description])
        pickle.dump(Calendar,open("Calendar.dat","wb"))
        Say("Event added to your calendar")

    try:
        HueInfo = pickle.load(open("PhilipsHue.dat","rb")) 
    except Exception:
        HueInfo = {}
        pickle.dump({},open("PhilipsHue.dat","wb")) 
    def SetupHue(cmd):
        Say("Welcome to Philips Hue light setup")
        try:
            Say("Retrieving your Hue bridge's ip")
            searchInput = "https://www.meethue.com/api/nupnp"
            res = requests.get(searchInput)
            res.raise_for_status()
            info = bs4.BeautifulSoup(res.text, "html.parser").text
            info = ast.literal_eval(info)[0]
            HueInfo["BridgeIp"] = info["internalipaddress"]
            
            Say("Creating a username")
            HueInfo["Username"] = qhue.create_new_username(HueInfo["BridgeIp"])
            
            Say("I would reccomend you turn on all of your Hue lights now. I will turn them off one at a time so you can give them a name to easily access them")
            b = qhue.Bridge(HueInfo["BridgeIp"],HueInfo["Username"])
            Say("Connected to your hue bridge")
            HueInfo["Lights"] = {}
            for light in b.lights():
                if b.lights[light].on:
                    b.lights[light].state(on = False)
                    lightName = SpeechInput("Find the light that I turned off. What name do you want to give this light: ").lower()
                    HueInfo["Lights"][lightName] = b.lights[light]()["uniqueid"]
            Say("Saving Hue Settings")
            pickle.dump(HueInfo,open("PhilipsHue.dat","wb")) 
            Say("Setup Complete")
        except Exception as e:
            Say("An error occured when setting up your Hue lights")

    def RenameHueLight(cmd):
        if HueInfo == {}:
            Say("Hmm, It appears you don't have this command set up")
            Say("To set up hue lights to work with me, simply say set up hue")
        Say("Here is a list of your lights: ")
        for light in HueInfo["Lights"]:
            Say(light)
        try:
            old = SpeechInput("Enter the name of the light").lower()
            HueInfo["Lights"][old]
        except:
            Say("I couldn't find a light called "+old)
            return
        new = SpeechInput("Enter the new name for the light").lower()
        HueInfo["Lights"][new] = HueInfo["Lights"][old]
        del HueInfo["Lights"][old]
        pickle.dump(HueInfo,open("PhilipsHue.dat","wb")) 
        Say("Renamed "+old+" to "+ new)
            
    def HuePercent(cmd):
        if HueInfo == {}:
            Say("Hmm, It appears you don't have this command set up")
            Say("To set up hue lights to work with me, simply say set up hue")
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
                Say("Turned "+lightName+" to "+str(cmd.cachedData[1])+"%")
            else:
                Say("I couldn't find a light called "+lightName)
        except:
            Say("Sorry, something went wrong")

    def HueOnOff(cmd):
        if HueInfo == {}:
            Say("Hmm, It appears you don't have this command set up")
            Say("To set up hue lights to work with me, simply say set up hue")
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
                Say("Turned "+ cmd.cachedData[1]+" "+lightName)
            else:
                Say("I couldn't find a light called "+lightName)
        except:
            Say("Sorry, something went wrong")

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
                    if "y" in SpeechInput("Would you like to remove this event from your calendar now?").lower():
                        Calendar.remove(event)
            pickle.dump(Calendar,open("Calendar.dat","wb"))
            if not found:
                Say("I couldn't find anything on your calendar for today")
        else:
            Say("There isn't any events in your calendar")

    def BasicMath(cmd):
        a = cmd.cachedData[0]
        b = cmd.cachedData[1]
        r = 0
        operation = cmd.cachedData[2]
        if operation in ["add","plus","+"]:
            r = a + b
        elif operation in ["minus","subtract","-"]:
            r = a - b
        elif operation in ["divide","over"]:
            r = a / b
        elif operation in ["multiply","times"]:
            r = a * b
        elif operation in ["power","^"]:
            operation = "to the power of"
            r = a ** b
        else:
            Say("I don't recognise that operation...")
            return
        try:
            if r % 1 == 0:
                r = int(r)
            else:
                round(r)
        except Exception:
            pass
        Say("The result of "+str(a)+" "+operation+" "+str(b)+ " is "+ str(r))
        

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
        Say(joke)
        st = time.time()
        Say(jokeBook[joke])


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
        appName = SpeechInput("What is the name of the app you would like to create or change?").lower()
        appName = "".join(FilterAndSeperate(appName,[" "])).lower()
        if appName in list(appDirs.keys()):
            if "y" in SpeechInput("An app called "+appName+" already exists, are you sure you want to change it?").lower():
                appDirs[appName] = GetDirectory()
                pickle.dump(appDirs,open("AppInfo.dat","wb"))
                Say("Changed app called"+appName+" successfully")
        else:
            if "y" in SpeechInput("Are you sure you want to add an app called "+appName).lower():
                appDirs[appName] = GetDirectory()
                pickle.dump(appDirs,open("AppInfo.dat","wb"))
                Say("Added app called"+appName+" successfully")

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
        Apps = c.Win32_Process()

        for p in Apps:
            if not p.ExecutablePath in list(appDirs.values()) and not p.ExecutablePath in ignoredDirs:
                if "y" in SpeechInput("I noticed an app I haven't seen before called "+ p.Name+". Would you like to create a shortcut to it?").lower():
                    appName = SpeechInput("What would you like to call it?")
                    appName = "".join(FilterAndSeperate(appName,[" "])).lower()
                    if appName in list(appDirs.keys()):
                        if "y" in SpeechInput("An app called "+appName+" already exists, are you sure you want to change it?").lower():
                            appDirs[appName] = p.ExecutablePath
                            pickle.dump(appDirs,open("AppInfo.dat","wb"))
                    else:
                        if "y" in SpeechInput("Are you sure you want to add an app called "+appName).lower():
                            appDirs[appName] = p.ExecutablePath
                            pickle.dump(appDirs,open("AppInfo.dat","wb"))
                else:
                    ignoredDirs.append(p.ExecutablePath)
                    pickle.dump(ignoredDirs,open("IgnoredDirectories.dat","wb")) 


    def OpenApp(cmd):
        global appDirs
        appName = cmd.cachedData[0]
        selectedApp = ""
        for app in appDirs:
            if app == appName:
                selectedApp = appDirs[app]
        if selectedApp != "":
            Say("Opening "+ appName)
            if not os.path.isfile(selectedApp):
                Say("The directory to "+ selectedApp + "is incorrect or has been changed")
                Say("You can now change it to the correct directory")
                appDirectory = GetDirectory()
                selectedApp = appDirectory
                appDirs[appName] = appDirectory
                pickle.dump(appDirs,open("AppInfo.dat","wb"))
            os.startfile(selectedApp)
        else:
            Say("I couldn't find an app called " + appName)
            if "y" in SpeechInput("Would you like to create an app called "+appName+"?: ").lower():
                appDirectory = GetDirectory()
                appDirs[appName] = appDirectory
                pickle.dump(appDirs,open("AppInfo.dat","wb"))
                Say("Successfully added "+ appName)
                os.startfile(selectedApp)
         
    def StartTimer(cmd):
        global startTime
        startTime = time.time()
        Say("Timer started!")

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
            Say("Timer Stopped:")
            Say("Hours: "+str(hours)+", Minutes: "+ str(minutes)+", Seconds: "+ str(seconds)+", Milliseconds: "+str(milliseconds))
        else:
            if "y" in SpeechInput("You never started a timer. Would you like to start a timer now?: ").lower():
                StartTimer(cmd)

    try:
        userLocation = pickle.load(open("location.dat","rb"))
    except Exception:
        while True:
            userLocation = SpeechInput("For tasks that require infomation local to where you live, what town do you live in?")
            if "y" in SpeechInput("Are you sure you want to set "+userLocation+" as your location?").lower():
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
                suggestion = "Its going to be freezing today so wrap up really warm "
            elif (temp < 15 and temp > 0) or status == "snow":
                suggestion= "Wrap up warm. It should be cold today"
            elif temp > 22 and temp <= 30:
                suggestion= "It should be nice and hot today"
            elif temp > 30:
                suggestion= "Its very hot today. Remember to drink lots of water to stay hydrated"
            else:
                suggestion = "The temperature should be quite mild"
            if status == "clouds":
                text += "cloudy"
            elif status == "mist":
                suggestion += " and be careful if you're driving"
                text += "misty"
            elif status == "rain" or status == "drizzle":
                suggestion += " and don't forget a coat and umbrella"
                text += "raining"
            elif status == "snow":
                suggestion += " and don't forget to wear a coat. Have fun in the snow!"
                text += "snowing"
            else:
                text += status
            text += " at a temperature of "+str(round(temp))+" degrees celsius. " + suggestion
            Say(text)
        except:
            Say("I am not sure what the weather is right now.")
        
        
        
            
        
        
    def Search(cmd):
        try:
            searchInput = cmd.cachedData[0]
            searchInput = "https://google.com/search?q="+searchInput
            res = requests.get(searchInput)
            res.raise_for_status()
            soup = bs4.BeautifulSoup(res.text, "html.parser")
            linkElements = soup.select('div#main > div > div > div > a')
            if len(linkElements) == 0:
                Say("Couldn't find any results...")
                return
            link = linkElements[0].get("href")
            i = 0
            while link[0:4] != "/url" or link[14:20] == "google":
                i += 1
                link = linkElements[i].get("href")
            Say("Searching google for " + cmd.cachedData[0])
            webbrowser.open("http://google.com"+link)
        except Exception as e:
            Say("Something went wrong while searching for "+ cmd.cachedData[0])

    try:
        shoppingList = pickle.load(open("shoppingList.dat","rb"))
    except:
        shoppingList = []
        
    def AddShoppingList(cmd):
        global shoppingList
        shoppingList.append(cmd.cachedData[1])
        pickle.dump(shoppingList,open("shoppingList.dat","wb"))
        Say("Added "+cmd.cachedData[1]+" to your shopping list")

        
    def RemoveShoppingList(cmd):
        global shoppingList
        try:
            shoppingList.remove(cmd.cachedData[1])
            pickle.dump(shoppingList,open("shoppingList.dat","wb"))
            Say("Removed "+cmd.cachedData[1]+" to your shopping list")
        except:
            pass

    def ClearShoppingList(cmd):
        global shoppingList
        try:
            shoppingList = []
            pickle.dump(shoppingList,open("shoppingList.dat","wb"))
            Say("Cleared your shopping list")
        except:
            pass

    def GoShopping(cmd):
        if shoppingList != []:
            try:
                for item in shoppingList:
                    searchInput = "https://google.com/search?q="+item+"amazon"
                    res = requests.get(searchInput)
                    res.raise_for_status()
                    soup = bs4.BeautifulSoup(res.text, "html.parser")
                    linkElements = soup.select('div#main > div > div > div > a')
                    if len(linkElements) == 0:
                        Say("Couldn't find any results...")
                    link = linkElements[0].get("href")
                    i = 0
                    while link[0:4] != "/url" or link[14:20] == "google":
                        i += 1
                        link = linkElements[i].get("href")
                    webbrowser.open("http://google.com"+link)
                if "y" in SpeechInput("Would you like to clear your shopping list?").lower():
                    ClearShoppingList("")
            except Exception as e:
                Say("You can't go shopping right now")
        else:
            Say("You don't have anything on your shopping list")

    try:
        email = pickle.load(open("email.dat","rb"))
    except:
        Say("Would you mind quickly setting up an email address that I can use to send you info")
        Say("Once you've created it, enter in its email and password so I can use it")
        email = {}
        Say("Type in my email address")
        email["Address"] = input(":")
        Say("Type in my password")
        email["Pass"] = input(":")
        Say("Type in your email address so I can send emails to you")
        email["UserAddress"] = input(":")
        pickle.dump(email,open("email.dat","wb"))

    def SendEmail(recipient,subject = "",message = "",attachment = ""):
        try:
            msg = MIMEMultipart()
            msg["From"] = email["Address"]
            msg["To"] = email["Address"]
            msg["Subject"] = subject
            if attachment != "":
                Attachment = open(attachment,"rb")
                part = MIMEBase("application","octet-stream")
                part.set_payload((Attachment).read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition","attachement; filename ="+attachment)
            msg.attach(MIMEText(message,"plain"))
            msg.attach(part)
            text = msg.as_string()
            server = smtplib.SMTP("smtp.gmail.com:587")
            server.ehlo()
            server.starttls()
            server.login(email["Address"],email["Pass"])
            server.sendmail(email["Address"],recipient,text)
            server.quit()
            Say("Email Sent")
        except Exception as e:
            app.DisplayOutput(e)
            Say("Email failed to send")

    def Email(cmd):
        Say("Type the email address of the email recipient")
        recipient = input(":")
        subject = SpeechInput("Enter a subject for the email")
        message = SpeechInput("Enter in a message")
        Say("Add an attachment:")
        attachment = GetDirectory()
        SendEmail(recipient,subject,message,attachment)

    def PrintShoppingList(cmd):
        try:
            List = docx.Document()
            List.add_heading("Shopping List:",1)
            i = 0

            for item in shoppingList:
                i += 1
                List.add_paragraph(str(i)+". "+item.capitalize())
            List.save("shoppingList.docx")

            word = comtypes.client.CreateObject('Word.Application')
            doc = word.Documents.Open(os.getcwd()+"\shoppingList.docx")
            doc.SaveAs(os.getcwd()+"\shoppingList.pdf", FileFormat=17)
            doc.Close()
            word.Quit()

            SendEmail(email["UserAddress"],"Shopping List","Here's your shopping list you can take with you","shoppingList.pdf")
        except:
            Say("Failed to send your shopping list")

    def NumericalFilter(text,validCharacters = ["0","1","2","3","4","5","6","7","8","9","-",".","+"]):
        result = ""
        for char in text:
            if char in validCharacters:
                result += char
        return result

    def Mute(cmd):
        global muted
        muted = True
        Say("muted")

    def Unmute(cmd):
        global muted
        muted = False
        Say("Unmuted")

    def Marsh(cmd):
        Say("Mr marsh is life")

    cmdWords = {}
    lastCMD = None
    lastFilterInput = []
    CMDs = {}
    def Help(cmd):
        while True:
            Say("Here are the available topics:")
            Say("Commands")
            Say("How I Work")
            Say("Talking")
            Topic = SpeechInput("What Topic Would you like me to cover?").lower()

            if Topic == "commands":
                Say("Here's a list of the commands you can get me to do: ")
                for cmd_ in CMDs:
                    Say(cmd_)
                commandHelp = True
                while commandHelp:
                    query = SpeechInput("What would you like help with? (say done to leave help)").lower()
                    if query == "done":
                        commandHelp = False
                    else:
                        try:
                            Say(CMDs[query].help)
                            Say("Required Keywords:")
                            app.DisplayOutput(str(CMDs[query].reqWords))
                            time.sleep(4)
                        except Exception:
                            Say("I couldn't help you with the query, "+query)
            elif Topic == "how i work" or Topic == "how you work":
                Say("I work by analysing each word you tell me")
                Say("I have a certain set of commands of which some have required keywords that I use to figure out what you want me to do.")
                Say("Some keywords help tell me if I need to be looking for infomation")
                Say("such as the open keyword that tells me the name of the app will come after it.")
                Say("This can help give you more flexibility in how you say certain commands as I only look for keywords.")
                Say("Although I can get a bit confused if you don't use the required keywords.")
                Say("I work out what you are saying by using Google's speech to text API which requires an internet connection.")
                Say("I can't work out what you say to me without a connection and so you will be limited to typing in commands.")
            elif Topic == "talking":
                Say("When talking to me you may notice several dots appearing in the console")
                Say("1 dot means I am listening for your command")
                Say("2 dots means I am working out what you said to me")
                Say("3 dots means I am ready to carry out your given command")
                Say("I can't hear you while I am working out or carrying out a command")
                Say("So make sure you only talk to me while there is 1 dot on screen")
                Say("Also, in some cases you might want to type what you want me to do.")
                Say("If so you can switch between the two modes by saying 'manual' to type and type 'speech' to start talking again.")
            else:
                Say(Topic+" is not a valid topic")
            if "n" in SpeechInput("Would you like to continue?").lower():
                break
        Say("Bye for now! Remember, I am always here to help you!")
                    
    #Command Handling

    class cmdWord:
        def __init__(self,word,startRecording = False,stopRecording = False,addCMDWord = False):
            self.word = word
            self.includeWord = addCMDWord
            self.start = startRecording
            self.stop = stopRecording
            cmdWords[word] = self

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
        def __init__(self,name,requiredCmdWords = [[""]],function = NoReply,helpDescription = "",dataFormat = [cmdType()]):
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
                GetData = False
                criteria = FindEmptyData(self,self.cachedData)
                if criteria == True:
                    break
                Data = ""
                for word in words:
                    if word in cmdWords.keys():
                        if cmdWords[word].includeWord:
                            cachedData = AddCachedData(word,self.cachedData,self.dataFormat)
                            criteria = FindEmptyData(self,self.cachedData)
                            if criteria == True:
                                break

                        if cmdWords[word].stop:
                            GetData = False
                            self.cachedData = AddCachedData(Data,self.cachedData,self.dataFormat)
                            Data = ""
                            criteria = FindEmptyData(self,self.cachedData)
                            if criteria == True:
                                break
                            
                    if GetData:
                        if Data != "" and criteria.multiWord:
                            Data += " "
                        if criteria.numerical:
                            Data += NumericalFilter(word)
                        else:
                            Data += word
                    if word in cmdWords.keys():
                        if cmdWords[word].start:
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
            if CMDs[cmd_].relevantCommand(Words) and len(CMDs[cmd_].reqWords) > len(CMDs[best].reqWords):
                best = cmd_
        return CMDs[best]

    def Power(cmd):
        if cmd.cachedData[0] == "shutdown":
            if "y" in SpeechInput("Are you sure you want to shutdown your computer?").lower():
                os.system("shutdown /s")
        elif cmd.cachedData[0] == "restart":
            if "y" in SpeechInput("Are you sure you want to restart your computer?").lower():
                os.system("shutdown /r")
        elif cmd.cachedData[0] == "logout":
            if "y" in SpeechInput("Are you sure you want to log out of your computer?").lower():
                os.system("shutdown /l")

    #CommandWords

    cmdWord("off",False,False,True)
    cmdWord("on",False,False,True)
    cmdWord("turn",True)
    cmdWord("light",False,True)
    cmdWord("shutdown",False,False,True)
    cmdWord("logout",False,False,True)
    cmdWord("restart",False,False,True)
    cmdWord("translate",True)
    cmdWord("and",True,True)
    cmdWord("to",True,True)
    cmdWord("from",True,True)
    cmdWord("by",True,True)
    cmdWord("open",True,False)
    cmdWord("start",True,False)
    cmdWord("search",True)
    cmdWord("number",True)
    cmdWord("total",True)
    cmdWord("result",True)
    cmdWord("add",True,True,True)
    cmdWord("remove",True,True,True)
    cmdWord("subtract",True,True,True)
    cmdWord("+",True,True,True)
    cmdWord("-",True,True,True)
    cmdWord("minus",True,True,True)
    cmdWord("power",True,True,True)
    cmdWord("^",True,True,True)
    cmdWord("minus",True,True,True)
    cmdWord("plus",True,True,True)
    cmdWord("divide",True,True,True)
    cmdWord("over",True,True,True)
    cmdWord("multiply",True,True,True)
    cmdWord("times",True,True,True)
    cmdWord("when", True)

    #Commands
    CMD("no reply",[],NoReply,"This just tells you that I didn't quite understand what you wanted me to do.",[])
    CMD("help",["help"],Help,"Provides help",[])
    CMD("repeat",[["repeat"]],Repeat,"Repeats the last command",[])
    CMD("hello",[["hello","hi","hey"]],Hello,"Simply say hello.",[])
    CMD("goodbye",[["goodbye","bye"]],Goodbye,"Simply say goodbye to close me.",[])
    CMD("change my name",[["change"],["your"],["name","username"]],ChangeName,"Change my name to something that feels more personal",[])
    CMD("change your name",[["change"],["my"],["name","username"]],ChangeUsername,"Change your username to something that feels more personal",[])
    CMD("change voice",[["change"],["voice"]],ChangeVoice,"Change my voice to one of several option that are demoed to you beforehand",[])
    CMD("open app",[["open"]],OpenApp,"Open an app by telling me its name. If I don't already know where that app is I'll ask you for its directory so I can create a shortcut",[cmdType(False,False)])
    CMD("change app",[["change","add"],["app"]],ChangeApp,"If you made a mistake when giving me an app's directory or want to add an new app's directory",[])
    CMD("google search",[["search", "google"]],Search,"Search google using the 'search' or 'google' keyword and I will provide you with the top search result")
    CMD("joke", [["joke"]],Joke,"I'll tell you a joke",[])
    CMD("time", [["time"]],GetTime,"I'll tell you the time",[])
    CMD("date",[["date","day","today"]],GetDate,"I'll tell you the date",[])
    CMD("random number",[["random"],["number"],["and","to"]],RandomNumber,"I'll give you a random number between a and b which could be set out similar to the following: give me a random number between a and b.",[cmdType(True,False),cmdType(True,False)])
    CMD("simple math",[["add","plus","subtract","minus","times","power","^","multiply","divide","over","+","-"],["result","total"]],BasicMath,"I can do division, multiplication, subtraction and addition of two numbers but remember to use the keywords 'result' or 'total'",[cmdType(True,False),cmdType(True,False),cmdType(False,False)])
    CMD("start timer",[["start","begin"],["timer"]],StartTimer,"I'll start a timer",[])
    CMD("stop timer",[["stop","end","finish"],["timer"]],StopTimer,"I'll stop a timer and tell you how long it lasted",[])
    CMD("events",[["when"],["is"]],WhenEvent,"I can tell you when special events like halloween are")
    CMD("add calendar",[["add"],["calendar"]],AddCalendar,"Set events on certain dates to remind you to do tasks and stay organised",[])
    CMD("check calendar",[["check","show"],["calendar"]],CheckCalendar,"Check tasks you have to do today on your calendar to remind you to do tasks and stay organised",[])
    CMD("???",[["mr"],["marsh"],["is"],["love"]],Marsh,"???",[])
    CMD("mute",[["mute"]],Mute,"Mute my voice but text will still appear on screen",[])
    CMD("unmute",[["unmute"]],Unmute,"Unmute my voice when you want",[])
    CMD("add shopping list",[["add"],["to"],["shopping"],["list"]],AddShoppingList,"Add something to your shopping list",[cmdType(False,False),cmdType(False,True)])
    CMD("remove shopping list",[["remove"],["to"],["shopping"],["list"]],RemoveShoppingList,"Remove an item from your shopping list",[cmdType(False,False),cmdType(False,True)])
    CMD("clear shopping list",[["clear"],["shopping"],["list"]],ClearShoppingList,"Removes all items from shopping your list",[])
    CMD("go shopping",[["go"],["shopping"]],GoShopping,"Searches each item on your shopping list on amazon",[])
    CMD("email shopping list",[["email","send","give","get","app.DisplayOutput"],["shopping"],["list"]],PrintShoppingList,"Emails your shopping list to the email that you provided me with",[])
    CMD("email",[["send","write"],["email"]],Email,"Sends an email through Gmail",[])
    CMD("translate",[["translate"],["from"],["to"]],Translate,"Translate some words from one language to another",[cmdType(False,True),cmdType(False,False),cmdType(False,False)])
    CMD("weather",[["weather"]],Weather,"I'll tell you waht the weather is like today and suggest what to wear",[])
    CMD("power",[["shutdown","logout","restart"],["computer"]],Power,"Allows you to restart/shutdown/logout simply by asking",[cmdType(False,False)])
    CMD("hue setup",[["set"],["up"],["hue","lights"]],SetupHue,"Allows you to use your Philips Hue lights with me",[])
    CMD("hue turn on off",[["turn"],["light"],["on","off"]],HueOnOff,"Turn a Philips Hue light on or off",[cmdType(False,True),cmdType(False,False)])
    CMD("rename hue light",[["rename"],["hue","light"]],RenameHueLight,"Rename a light to something that better suits it. This name is used to say what light should be turned on/off etc.",[])
    CMD("hue light brightness",[["turn"],["light"],["to"]],HuePercent,"Set the brightness of a given light as a given percent",[cmdType(False,True),cmdType(True,False)])
    CMD("flip coin",[["flip"],["coin"]],FlipCoin,"I'll flip a coin for you incase you want to settle an argument",[])
    CMD("roll dice",[["roll"],["die","dice"]],RollDice,"I can roll a die if you don't have one",[])
    CMD("restart",[["restart","assistant"]],Restart,"Restart me",[])

    #40 commands

    CheckOpenApps()
    lastCheck = time.time()

    try:
        username = pickle.load(open("Username.dat","rb"))
        Hello("")
        Say("Simply ask for help if you need anything!")
    except Exception:
        Say("Hello I don't believe we've met before!")
        username = SpeechInput("What's your name?")
        Say("Hello, "+username+"! That's a nice name. my name is "+ name+ " but you can change it if you want.")
        pickle.dump(username, open("Username.dat","wb"))
        Say("I recommend you start off in help to see what you can do")
        Help("")

    appCheckFrequency = 3 * 60

    while True:
        if time.time() - lastCheck >= appCheckFrequency:
            CheckOpenApps()
            lastCheck = time.time()
        app.dots.configure(text = ".")
        if not manualMode or (manualMode and app.submitted):
            userInput = SpeechInput("",2).lower()
            if userInput != "" and userInput != "speech" and userInput != "manual":
                filterInput = FilterAndSeperate(userInput)
                UserCMD = BestCMD(filterInput)
                UserCMD.run(filterInput)
                lastFilterInput = filterInput
                if UserCMD.function != Repeat:
                    lastCMD = UserCMD
        if windowClosed:
            exit()
                        
except BaseException as e:
    Say("An error occured")
    app.dots.configure(text = "ERROR: "+str(e))
    Say("Automatically restarting in 5 seconds")
    time.sleep(5)
    Restart("")


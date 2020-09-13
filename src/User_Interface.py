from tkinter import *
from tkinter import ttk
import drafter_VOEs
import config
import data_analytics
import drafter_JDReqs
import drafter_ETA9089Review
import drafter_PERMsFiled

###This is the main UI module
### the UI is Object Oriented

class JeromesButtons:

    def __init__(self, master):
        frame = Frame(master)
        frame.grid()

        ### Entry field
        self.label_1 = Label(frame, text="Please enter the relevant fields:")
        self.label_1.grid(row = 0, column=0, sticky = W + E, pady = 5)
        self.ent = Entry(frame, width = 22)
        self.ent.grid(row = 0, column = 1, padx = 5, pady = 10, sticky = W)

        ### Combo Field
        self.options = ["Auto-Draft JD&Reqs (Format is C03:Q1000)", "Auto-Draft VOE's (Format is C03:R1000)",
                        "Auto-Draft ETA9089 for EE Review (Format is C03:AA1000)", "Auto-Draft PERM Filed (Format is C03:AG1000)"]
        self.combofield = ttk.Combobox(frame, value = self.options, width = 50)
        self.combofield.grid(row = 1, column = 0, columnspan = 2, padx = 10, pady = 5)
        self.combofield.current(0)
        self.combofield.bind("<<ComboboxSelected>>", self.selection)

        ###### Information Output for the end-user
        ### Labelframe field. Should be able to display relevant information to the user.
        self.Labelframe = LabelFrame(frame, bd = 2, bg = 'white', text = "PERM Automation Framework")
        self.Labelframe.grid(row = 2, column = 0, columnspan = 2, pady = 5, padx =30)
        ### Label
        self.label = Label(self.Labelframe, text = "Overview", bg = 'white', justify = CENTER)
        self.label.pack(anchor = W)
        ### Scrollbar
        self.scrollbar = Scrollbar(self.Labelframe)
        self.scrollbar.pack( side = RIGHT, fill = Y)
        ### Listbox within Labelframe
        self.listbox = Listbox(self.Labelframe, bg = 'white', yscrollcommand = self.scrollbar.set, width = 50, height = 20)
        self.listbox.pack(side = LEFT, fill = BOTH)
        self.scrollbar.config(command = self.listbox.yview)
        ### Clear listbox button
        self.clear = Button(frame, text = "Clear/Refresh", padx = 10, pady = 5, command = self.clear)
        self.clear.grid(row = 3, column = 0, columnspan = 1, sticky = W, padx = 30)
        ###
        #################
        ### Analytics Button
        self.analytics = Button(frame, text = "Analytics Dashboard", padx = 10, pady = 5, command = data_analytics.mainanalytics)
        self.analytics.grid(row = 3, column = 1, sticky = E, padx = 30)

        ### this is the run function button. it does not run the function but rather accepts an input that will be passed into the gatherdata module.
        self.runfunction = Button(frame, text = "Run/Execute", padx = 50, pady = 10, command = self.func)
        self.runfunction.grid(row = 4, column = 0, pady = 20, padx = 30, sticky = W)

        ### quit button which is rather self explanatory
        self.quitButton = Button(frame, text = "Exit", padx = 75, pady = 10, command = frame.master.destroy)
        self.quitButton.grid(row = 4, column = 1, sticky = E, pady = 20, padx = 30)

    def selection(self, event):
        select = self.combofield.get()

    def clear(self):
        self.listbox.delete(0, 'end')

    ###this function returns the input from the entry field. It turns the input from the entry field into a workable object.
    def func(self):
        select = self.combofield.get()
        ###############
        if select == "Auto-Draft JD&Reqs (Format is C03:Q1000)":
            config.fields = self.ent.get()
            #### Executing the main functionality of the application (gathering data from spreadsheet + drafting)
            data_analytics.mainloopJDReqs()
            drafter_JDReqs.maindraft()
            #########
            ######## outputting the information into the listbox/labelframe
            length = len(config.information)
            analytics = "Drafted {number} JD&Reqs emails.".format(number=length)
            self.listbox.insert(END, analytics)
            ######## Emails drafted for the following indivudals
            self.listbox.insert(END, "Emails drafted for the following individuals:")
            counter = 0
            for i in config.information:
                self.listbox.insert(END, "{LAST}, {First}, Option {option}".format(LAST= config.information[counter][0],
                                                                                   First= config.information[counter][1],
                                                                                   option= config.information[counter][9]))
                counter = counter + 1
            self.listbox.insert(END, "\n")
        ###############
        ###############
        if select == "Auto-Draft VOE's (Format is C03:R1000)":
            config.fields = self.ent.get()
            #### Executing the main functionality of the application (gathering data from spreadsheet + drafting)
            data_analytics.mainloopVOE()
            drafter_VOEs.mainfunc()
            #########
            ######## outputting the information into the listbox/labelframe
            length = len(config.information)
            analytics = "Drafted {number} VOE emails.".format(number = length)
            self.listbox.insert(END, analytics)
            ######## Emails drafted for the following indivudals
            self.listbox.insert(END, "Emails drafted for the following individuals:")
            counter = 0
            for i in config.information:
                self.listbox.insert (END, "{LAST}, {First}, Option {option}".format(LAST = config.information[counter][0],
                                                                                    First = config.information[counter][1],
                                                                                    option = config.information[counter][6]))
                counter = counter + 1
            self.listbox.insert(END, "\n")
        ###############
        ###############
        if select == "Auto-Draft ETA9089 for EE Review (Format is C03:AA1000)":
            config.fields = self.ent.get()
            #### Executing the main functionality of the application (gathering data from spreadsheet + drafting)
            data_analytics.mainloopETAReview()
            drafter_ETA9089Review.ETA9089draft()
            #########
            ######## outputting the information into the listbox/labelframe
            length = len(config.information)
            analytics = "Drafted {number} ETA9089 EE Review emails.".format(number=length)
            self.listbox.insert(END, analytics)
            ######## Emails drafted for the following indivudals
            self.listbox.insert(END, "Emails drafted for the following individuals:")
            counter = 0
            for i in config.information:
                self.listbox.insert(END, "{LAST}, {First}".format(LAST=config.information[counter][0],
                                                                First=config.information[counter][1]))

                counter = counter + 1
            self.listbox.insert(END, "\n")
        ###############
        ###############
        elif select == "Auto-Draft PERM Filed (Format is C03:AG1000)":
            config.fields = self.ent.get()
            #### Executing the main functionality of the application (gathering data from spreadsheet + drafting)
            data_analytics.mainloopPERMFiled()
            drafter_PERMsFiled.mainfunction()
            #########
            ######## outputting the information into the listbox/labelframe
            length = len(config.information)
            analytics = "Drafted {number} PERM's Filed emails.".format(number=length)
            ######## Emails drafted for the following indivudals
            self.listbox.insert(END, analytics)
            self.listbox.insert(END, "Emails drafted for the following individuals:")
            counter = 0
            for i in config.information:
                self.listbox.insert(END, "{LAST}, {First}".format(LAST=config.information[counter][0],
                                                                First=config.information[counter][1]))
                counter = counter + 1
            self.listbox.insert(END, "\n")
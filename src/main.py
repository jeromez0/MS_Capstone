import User_Interface
from tkinter import *
from PIL import ImageTk, Image
from tkinter import messagebox
import config

login = {"username": "jeromezhang", "password":"jerome"}

#### user interface for the login screen
class loginscreen:
    #######
    def __init__(self, master):
        self.frame = Frame(master)
        self.frame.grid()
        #####username label
        self.labelusername = Label(self.frame, text = "Username:")
        self.labelusername.grid(row=0, column=0, padx=5, pady=10)
        #####password label
        self.labelpassword = Label(self.frame, text = "Password:")
        self.labelpassword.grid(row=1, column=0, padx=5, pady=10)
        ##### username entry field
        self.entryusername = Entry(self.frame, width = 25)
        self.entryusername.grid(row=0, column=1, padx=5, columnspan=2)
        ##### password entry field
        self.entrypassword = Entry(self.frame, width = 25, show = "*")
        self.entrypassword.grid(row=1, column=1, padx=5, columnspan=2)
        ##### JZ logo image
        self.my_img = ImageTk.PhotoImage(Image.open("images/jzlogo.png"))
        self.mylabel = Label(self.frame, image = self.my_img)
        self.mylabel.grid(row = 2, column = 0, columnspan = 3, padx = 20)
        ##### login button
        self.loginbutton = Button(self.frame, text = "Login", padx = 60, command = self.enter)
        self.loginbutton.grid(row=3, column=0, columnspan=2, padx=5, pady=20)
        ##### exit button
        self.exitbutton = Button(self.frame, text = "Exit", padx= 30, command = self.frame.master.destroy)
        self.exitbutton.grid(row = 3, column = 2)
    #######
    def peasant(self):
        message = messagebox.showerror("incorrect pass", "try again" )

    #### this is the function that determines if the user logs in or backs out
    def enter(self):
        ### setting the variables for the login screen, log and passw
        log = self.entryusername.get()
        passw = self.entrypassword.get()
        ### is the login is the same as the designated username
        if log == login["username"]:
            ### if the password and the login match the dictionary values
            if passw == login["password"]:
                self.frame.master.destroy()     ### close the login screen
                main()                          ### start up the main program
            ### if the login and password don't match up, then open the message box
            else:
                self.peasant()
                config.count = config.count + 1
                if config.count == 5:       ### you get 5 chances to login otherwise the app will close
                    self.frame.master.destroy()
        ### if the login and password don't match up, then open the message box
        else:
            self.peasant()
            config.count = config.count + 1
            ### you get 5 chances to login otherwise the app will close
            if config.count == 5:
                self.frame.master.destroy()
    ### function that opens the messagebox

###### the main function runs the whole automation framework
def main():
    root = Tk()
    stuff = User_Interface.JeromesButtons(root)
    root.title("Generics PERM Automation Framework")
    root.mainloop()
###### the initial screen runs the login screen; to access the automation framework you must login
def initialscreen():
    root = Tk()
    sauce = loginscreen(root)
    root.title("Generics PERM Automation Framework")
    root.mainloop()
###### runs the whole thang
if __name__ == "__main__":
    initialscreen()
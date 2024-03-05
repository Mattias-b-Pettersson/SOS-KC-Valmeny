from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchWindowException
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from subprocess import CREATE_NO_WINDOW
from time import sleep
from datetime import datetime
from tkinter import messagebox, VERTICAL, ttk
from pythoncom import CoInitialize, CoUninitialize
from sys import exit
import win32com.client as win32
import threading
import customtkinter
import os

#######################################################
# Skapad av Mattias Pettersson @ Serviceförvaltningen #
#######################################################

prod = False
log_string = ""
first_name = ""
last_name = ""
times_run = 0
global driver
global action

kk_obstetriken_enheter = "Antenatalavdelning, Vårdavdelning 63 BB, Vårdavdelning 73 BB, Vårdavdelning 85 BB, Mödrahälsovårdsenhet, Förlossningsavdelning, \
    Ultraljudsmottagning, Amningscentrum, BB-mottagning"

kk_gynekologen_enheter = "Vårdavdelning 72, Akutmottagning för våldtagna, Gynekologisk mottagning, Gynekologisk akutmottagning, Modul 5 Gynekologi" 


class NoUserFoundException(Exception):
    """Raised when no user is found"""
    pass


class LifeCareAccountInactive(Exception):
    """Raised user account is inactive in LifeCare"""
    pass


class MoreThanOneAvailablePersonpost(Exception):
    """Raised when more than one available personpost"""
    pass


def get_input():
    """
    Main funktionen som bygger upp UI:n med biblioteket tkinter och skickar vidare informationen i UIn till rätt funktioner.  
    """

    # workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account
    global text_box
    global root

    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("green")

    root = customtkinter.CTk()
    root.title("SÖS-KC Valmeny")
    root.resizable(False, False)



    created_by = customtkinter.CTkLabel(
        master=root,
        text="Skapad av Mattias Pettersson @ Serviceförvaltningen",
        font=("Roboto", 12),
        )
    created_by.place(
        relx = 0.0,
        rely = 1.0,
        anchor ='sw'
    )
    created_by.grid(row=1, column=0)

    main_label = customtkinter.CTkLabel(master=root, text="SÖS-KC Valmeny", font=("Roboto", 32))
    main_label.grid(row=0, column=0, columnspan=2, pady=(30), padx=60)

    style = ttk.Style()
    style.configure("TNotebook", background="#242424")
    style.configure("TNotebook.Tab", padding=[10, 5], font=("Helvetica", 12), foreground="black")
    tabview = customtkinter.CTkTabview(root)
    tabview.grid(row=1, column=0, pady=(0,20), padx=60)

    
    
# -------- Nedan är Tab 1
    tabview.add("Ny profil")
    
    hsa_id_label = customtkinter.CTkLabel(master=tabview.tab("Ny profil"), text="Ange ett HSA-ID:", font=("Roboto", 14))
    hsa_id_label.grid(row=1, column=0, pady=(10,0), padx=(30,10))
    hsa_id_entry = customtkinter.CTkEntry(master=tabview.tab("Ny profil"), placeholder_text="HSA-ID")
    hsa_id_entry.grid(row=2, column=0, pady=(0,20), padx=(30,10))
    
    titel_label = customtkinter.CTkLabel(master=tabview.tab("Ny profil"), text="Vilken titel har användaren?", font=("Roboto", 14))
    titel_label.grid(row=3, column=0,pady=0, padx=(30,10))
    titel_entry = customtkinter.CTkComboBox(
        master=tabview.tab("Ny profil"),
        state="readonly",
        values=[
            "Läkare",
            "AT-Läkare",
            "Sjuksköterska",
            "Undersköterska",
            "Paramedicinerare",
            "Barnmorska",
            "Läkarsekreterare",
            "Student sjuksköterska",
            "Student paramedicinerare",
        ],
    )

    titel_entry.grid(row=4, column=0, pady=(0,20), padx=(30,10))

    workplace_label = customtkinter.CTkLabel(master=tabview.tab("Ny profil"), text="Vart jobbar användaren?", font=("Roboto", 14))
    workplace_label.grid(row=5, column=0, pady=0, padx=(30,10))
    workplace_entry = customtkinter.CTkComboBox(
        master=tabview.tab("Ny profil"),
        state="readonly",
        values=[
            "Akuten",
            "Perioperativ vård och Intensivvård - PVI",
            "Bild", 
            "Infektionskliniken-Venhälsan",
            "Internmedicin",
            "Kardiologin",
            "Kirurgin",
            "Ortopedin",
            "Urologin",
            "Sachsska Barnsjukhuset",
            "Onkologiska kliniken",
            "Ögon",
            "Handkirurgin",
            "Kvinnosjukvård och förlossning",
            "Obstetriken",
            "Gynekologen",
            "Sjukhusgemensam personal (AT-Läkare)",
        ],
    )
    workplace_entry.grid(row=6, column=0, pady=0, padx=(30,10))

    ids_entry = customtkinter.CTkSwitch(master=tabview.tab("Ny profil"), text="Skicka mail för IDS?", onvalue=True, offvalue=False)
    ids_entry.grid(row=7, column=0, pady=0, padx=(30,10))

    case_number_label = customtkinter.CTkLabel(master=tabview.tab("Ny profil"), text="Ange ärendenummret:", font=("Roboto", 14))
    case_number_label.grid(row=8, column=0, pady=0, padx=(30,10))
    case_number_entry = customtkinter.CTkEntry(master=tabview.tab("Ny profil"), placeholder_text="Ärendenummer")
    case_number_entry.grid(row=9, column=0, pady=(0,20), padx=(30,10))
    
    run_button = customtkinter.CTkButton(
        master=tabview.tab("Ny profil"), 
        text="Lägg på behörigheter", 
        command=lambda: run_button_func(
            workplace_entry.get(),
            titel_entry.get(),
            hsa_id_entry.get(),
            ids_entry.get(),
            run_button,
            loading_message,
            case_number_entry.get(),
            case_number_entry,
            hsa_id_entry,
        )
    )
    run_button.grid(row=10, column=0, pady=(0,20), padx=(30,10))

    create_newaccount_closemail_button = customtkinter.CTkButton(
        master=tabview.tab("Ny profil"), 
        text="Skapa standard mail \ntill nytt konto", 
        command=lambda: create_closemail_button_func(
            titel_entry.get(),
            hsa_id_entry.get(),
            case_number_entry.get(),
            case_number_entry,
            hsa_id_entry,
            new_or_remove="new"
        )
    )
    create_newaccount_closemail_button.grid(row=11, column=0, pady=(0,20), padx=(30,10))



# -------- Slut på  Tab 1
# # -------- Nedan är Tab 2
    tabview.add("Borttag")

    

    tab_2_hsa_id_label = customtkinter.CTkLabel(master=tabview.tab("Borttag"), text="Ange ett HSA-ID:", font=("Roboto", 14))
    tab_2_hsa_id_label.grid(row=0, column=0, pady=(10,0), padx=(30,10))
    tab_2_hsa_id_entry = customtkinter.CTkEntry(master=tabview.tab("Borttag"), placeholder_text="HSA-ID")
    tab_2_hsa_id_entry.grid(row=1, column=0, pady=(0,20), padx=(30,10))
    
    tab_2_titel_label = customtkinter.CTkLabel(master=tabview.tab("Borttag"), text="Vilken titel har användaren?", font=("Roboto", 14))
    tab_2_titel_label.grid(row=2, column=0,pady=0, padx=(30,10))
    tab_2_titel_entry = customtkinter.CTkComboBox(
        master=tabview.tab("Borttag"),
        state="readonly",
        values=[
            "Läkare",
            "AT-Läkare",
            "Sjuksköterska",
            "Undersköterska",
            "Paramedicinerare",
            "Barnmorska",
            "Läkarsekreterare",
            "Student sjuksköterska",
            "Student paramedicinerare",
        ],
    )
    tab_2_titel_entry.grid(row=3, column=0, pady=(0,20), padx=(30,10))

    tab_2_ids_entry = customtkinter.CTkSwitch(master=tabview.tab("Borttag"), text="Skicka mail för IDS?", onvalue=True, offvalue=False)
    tab_2_ids_entry.grid(row=4, column=0, pady=0, padx=(30,10))

    tab_2_case_number_label = customtkinter.CTkLabel(master=tabview.tab("Borttag"), text="Ange ärendenummret:", font=("Roboto", 14))
    tab_2_case_number_label.grid(row=5, column=0, pady=0, padx=(30,10))
    tab_2_case_number_entry = customtkinter.CTkEntry(master=tabview.tab("Borttag"), placeholder_text="Ärendenummer")
    tab_2_case_number_entry.grid(row=6, column=0, pady=(0,20), padx=(30,10))

    run_remove_button = customtkinter.CTkButton(
        master=tabview.tab("Borttag"), 
        text="Ta bort behörigheter", 
        command=lambda: run_remove_button_func(
            tab_2_titel_entry.get(),
            tab_2_hsa_id_entry.get(),
            run_remove_button,
            loading_message,
            tab_2_ids_entry.get(),
            tab_2_case_number_entry.get(),
            tab_2_case_number_entry,
            tab_2_hsa_id_entry,
        ) 
    )
    run_remove_button.grid(row=7, column=0, pady=(0,25), padx=(30,10))

    create_removeaccount_closemail_button = customtkinter.CTkButton(
        master=tabview.tab("Borttag"), 
        text="Skapa standard mail \ntill borttag av konto", 
        command=lambda: create_closemail_button_func(
            tab_2_titel_entry.get(),
            tab_2_hsa_id_entry.get(),
            tab_2_case_number_entry.get(),
            tab_2_case_number_entry,
            tab_2_hsa_id_entry,
            "remove"
        ) 
    )
    create_removeaccount_closemail_button.grid(row=8, column=0, pady=(0,25), padx=(30,10))


# -------- Slut på Tab 2    

    text_box = customtkinter.CTkTextbox(master=root, activate_scrollbars=True, width=800, height=400, font=("Roboto", 18))
    text_box.grid(row= 1, rowspan=10, column=1, pady=(20,20), padx=(0,30))
    text_box.configure(state="disabled")

    loading_message = customtkinter.CTkLabel(master=root, font=("Roboto", 14), text="", width=600)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    # notebook.select(frame_tab1)
    root.mainloop()

def create_closemail_button_func(user_titel, hsa_id_input, case_number, case_number_entry, hsa_id_entry, new_or_remove):
    """
    Funktion för att skapa ett nytt stängningsmail.
    """

    hsa_id_input = hsa_id_input.upper().replace(" ", "")

    if len(hsa_id_input) != 4:
        messagebox.showerror("SÖS-KC Valmeny error", "HSA-ID inte 4 karaktärer långt. Gör om gör rätt")
        return
    
    if len(case_number) == 0:
        messagebox.showerror("SÖS-KC Valmeny error", "Saknas ärendenummer")
        return

    # Skapar upp mailet med funktionen nedan
    if new_or_remove == "new":
        create_close_mail(user_titel, hsa_id_input, case_number, new_or_remove="new")
    elif new_or_remove == "remove":
        create_close_mail(user_titel, hsa_id_input, case_number, "remove")
    else:
        print("\nNågot gick fel... försök igen")
        root.after(0, print_text_in_text_box, f"\nNågot gick fel... försök igen")
        root.after(0, print_text_in_text_box, "-"*90 + "\n")
        
    # Rensar HSA-ID och äredenummer
    case_number_entry.delete(0, "end")
    hsa_id_entry.delete(0, "end")

    print("\nKlart. Väntar på ny input...")
    root.after(0, print_text_in_text_box, f"\n{hsa_id_input} Klar!")
    root.after(0, print_text_in_text_box, "-"*90 + "\n")

def check_input(hsa_id_input, user_titel, workplace_input, case_number, ids_account):
    """
    Dubbelkollar inputten som knapparna använder sig av.
    """

    continue_running = True
    hsa_id_input = hsa_id_input.upper().replace(" ", "")

    user_titel_full = ""
    user_titel_full += user_titel

    if user_titel == "Läkare":
        user_titel = "lak"
    elif user_titel == "AT-Läkare":
        user_titel = "at_lak"
    elif user_titel == "Sjuksköterska":
        user_titel = "ssk"
    elif user_titel == "Undersköterska":
        user_titel = "usk"
    elif user_titel == "Paramedicinerare":
        user_titel = "paramed"
    elif user_titel == "Barnmorska":
        user_titel = "barnmorsk"
    elif user_titel == "Läkarsekreterare":
        user_titel = "läksek"
    elif user_titel == "Student paramedicinerare":
        user_titel = "stud_paramed"
    elif user_titel == "Student sjuksköterska":
        user_titel = "stud_ssk"

    if len(hsa_id_input) != 4:
        messagebox.showerror("SÖS-KC Valmeny error", "HSA-ID inte 4 karaktärer långt. Gör om gör rätt")
        continue_running = False
    
    if user_titel == "":
        messagebox.showerror("SÖS-KC Valmeny error", "Ingen titel vald. Gör om gör rätt")
        continue_running = False
    
    if workplace_input == "":
        messagebox.showerror("SÖS-KC Valmeny error", "Inget VO valt. Gör om gör rätt")
        continue_running = False
    
    if case_number == "":
        make_sure_no_case_number = messagebox.askyesno("Är du säker?", "Är du säker på att du inte vill skapa ett lösningsmail?\n"
                                                       "Ärendenummer inte ifyllt")
        if make_sure_no_case_number:
            root.after(10, print_text_in_text_box, "Skippar att skapa ett lösningsmail")
        elif not make_sure_no_case_number:
            continue_running = False
        else:
            pass

    if not ids_account:
        if not workplace_input == "Obstetriken" and not user_titel == "Barnmorska":
            make_sure_ids = messagebox.askyesno("Är du säker?", "Är du säker på att du inte vill skicka ett mail för IDS?")
            if not make_sure_ids:
                continue_running = False
            elif make_sure_ids:
                root.after(10, print_text_in_text_box, "Skippar att skicka IDS mail...")
    
    if not continue_running: 
        return False, hsa_id_input, user_titel, user_titel_full
    elif continue_running:
        return True, hsa_id_input, user_titel, user_titel_full

def run_button_func(workplace_input, user_titel, hsa_id_input, new_ids_account, run_button, loading_message, case_number, case_number_entry, hsa_id_entry):
    """
    funktionen som knappen 'Lägg på behörigheter'/run knappen använder sig av. 
    Den kontrollerar så att inputten är korrekt angiven och sedan skickar vidare informationen till 'Handle input'
    """
    continue_running, hsa_id_input, user_titel, user_titel_full = check_input(hsa_id_input, user_titel, workplace_input, case_number, new_ids_account)

    if continue_running:
        pass
    elif not continue_running:
        return
    else: 
        root.after(10, print_text_in_text_box, "Något gick fel! Felanmäl detta! Felmeddelande: \nKunde inte kontrollera input.")
    

    if workplace_input == "Obstetriken" and user_titel == "Barnmorska":
        new_ids_account = False

    run_button.configure(state="disabled", bg_color="grey")

    loading_message.configure(text="Jobbar på behörigheterna...")
    loading_message.grid(row= 12, column=1, pady=(0,20), padx=30)

    workplace, vard_och_behandling_vmu_hsa, user_titel = handle_input(workplace_input, user_titel, hsa_id_input, new_ids_account, case_number)

    t = threading.Thread(target=run, args=(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, user_titel_full, hsa_id_input, new_ids_account, case_number))
    t.start()
    schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry)

def run_remove_button_func(user_titel, hsa_id_input, run_button, loading_message, send_ids_removal_mail, case_number, case_number_entry, hsa_id_entry):
    """
    funktionen som knappen 'ta bort behörigheter'knappen använder sig av. 
    Den kontrollerar så att inputten är korrekt angiven och sedan skickar vidare informationen till 'Handle input'
    """
    continue_running, hsa_id_input, user_titel, user_titel_full = check_input(hsa_id_input, user_titel, "no_new_workplace_input", case_number, send_ids_removal_mail)

    if continue_running == True:
        pass
    elif not continue_running:
        return
    else: 
        root.after(10, print_text_in_text_box, "Något gick fel! Felanmäl detta! Felmeddelande: \nKunde inte kontrollera input.")


    run_button.configure(state="disabled", bg_color="grey")

    loading_message.configure(text="Jobbar på borttagen...")
    loading_message.grid(row= 12, column=1, pady=(0,20), padx=30)

    t = threading.Thread(target=run_remove, args=(user_titel, user_titel_full, hsa_id_input, case_number, send_ids_removal_mail))
    t.start()
    schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry)

def schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry):
    """
    Schedule the execution of the `check_if_done()` function after
    500ms(0,5sek).
    """
    global root
    root.after(500, check_if_done, t, run_button, loading_message, case_number_entry, hsa_id_entry)


def check_if_done(t, run_button, loading_message, case_number_entry, hsa_id_entry):
    """
    Kollar om tråden är klar om den är klar. Återställer knappen och rensar fritext fälten.
    Tar även bort texten för laddningsmeddelandet 
    """

    if not t.is_alive():
        run_button.configure(state="normal", bg_color="green")
        loading_message.grid_forget()
        case_number_entry.delete(0, "end")
        hsa_id_entry.delete(0, "end")

    else:
        # Otherwise check again after one second.
        schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry)

def print_text_in_text_box(text):
    """
    Funktion för att printa text. Både i tkinter och i konsolen.
    """
    global root
    print(text)

    # Låser upp textboxen för att den är låst. Skriver in texten. Låser sedan textboxen igen. 
    text_box.configure(state="normal")
    text_box.insert(customtkinter.INSERT, f"{text}\n")
    text_box.configure(state="disabled")
    text_box.see(customtkinter.END)


def on_closing():
    """
    Denna funktionen specificerar vad som ska hända när man kryssar tkinter fönstret  
    """
    global root
    try:   
        driver.close()
    except:
        pass
    exit()
        

def handle_input(workplace_input, user_titel, hsa_id_input, new_ids_account, case_number):

    global handle_of_the_window_before_minimizing

    if user_titel == "Läkare":
        user_titel = "lak"
    elif user_titel == "AT-Läkare":
        user_titel = "at_lak"
    elif user_titel == "Sjuksköterska":
        user_titel = "ssk"
    elif user_titel == "Undersköterska":
        user_titel = "usk"
    elif user_titel == "Paramedicinerare":
        user_titel = "paramed"
    elif user_titel == "Barnmorska":
        user_titel = "barnmorsk"
    elif user_titel == "Läkarsekreterare":
        user_titel = "läksek"
    elif user_titel == "Student paramedicinerare":
        user_titel = "stud_paramed"
    elif user_titel == "Student sjuksköterska":
        user_titel = "stud_ssk"
    

    workplace_dictonary = {
        "Akuten": "aku",
        "Perioperativ vård och Intensivvård - PVI": "ane",
        "Bild": "bild",
        "Internmedicin": "int",
        "Kardiologin": "kar",
        "Kirurgin": "kir",
        "Ortopedin": "ort",
        "Urologin": "uro",
        "Sachsska Barnsjukhuset": "sac",
        "Onkologiska kliniken": "onk",
        "Ögon": "ögon",
        "Infektionskliniken-Venhälsan": "inf",
        "Handkirurgin": "hki",
        "Sjukhusgemensam personal (AT-Läkare)": "sgp",
        "Kvinnosjukvård och förlossning": "kk",
        "Gynekologen": "gyn",
        "Obstetriken": "obst",
    }

    workplace_input = workplace_dictonary[workplace_input]

    all_units_ek = {
        "aku": "Akut, Södersjukhuset AB",
        "ane": "Perioperativ vård och Intensivvård - PVI, Södersjukhuset AB",
        "bild": "Bilddiagnostik, Södersjukhuset AB",
        "int": "Internmedicin, Södersjukhuset AB",
        "kar": "Kardiologi, Södersjukhuset AB",
        "kir": "Kirurgi, Södersjukhuset AB",
        "ort": "Ortopedi, Södersjukhuset AB",
        "uro": "Urologiska kliniken, Specialistvård, Södersjukhuset AB",
        "sac": "Sachsska barn- och ungdomssjukhuset, Södersjukhuset AB",
        "kvi": "Kvinnosjukvård och förlossning, Södersjukhuset AB",
        "onk": "Onkologiska kliniken, Specialistvård, Södersjukhuset AB",
        "ögon": "Ögonklinik, Specialistvård, Södersjukhuset AB",
        "inf": "Infektionskliniken-Venhälsan, Specialistvård, Södersjukhuset AB",
        "hki": "Handkirurgi, Specialistvård, Södersjukhuset AB",
        "sgp": "Sjukhusgemensam personal, Södersjukhuset AB",
        "kk": "Kvinnosjukvård och förlossning, Södersjukhuset AB",
        "gyn": "Kvinnosjukvård och förlossning, Södersjukhuset AB",
        "obst": "Kvinnosjukvård och förlossning, Södersjukhuset AB",
    }

    workplace = all_units_ek[workplace_input]

    all_units_vmu = {
        "aku": "73N6",
        "ane": "7T5P",
        "bild": "7T5Q",
        "int": "73G7",
        "kar": "73MS",
        "kir": "73MR",
        "ort": "71XF",
        "uro": "7Q8W",
        "sac": "7VR0",
        "onk": "BGKZ",
        "ögon": "73N9",
        "inf": "7TKV",
        "hki": "7TKK",
        "gyn": "83L0",
        "obst": "83KZ",
        "sgp": "placeholder",
        "kk": "kk",
    }

    vard_och_behandling_vmu_hsa = all_units_vmu[workplace_input]

    return workplace, vard_och_behandling_vmu_hsa, user_titel

def open_browser():
    """
    Funktionen som körs för att öppna Edge
    """
    # öppnar edge
    ser = EdgeService(EdgeChromiumDriverManager().install())
    ser.creation_flags = CREATE_NO_WINDOW
    WINDOW_SIZE = "1920,1080"
    op = webdriver.EdgeOptions()
    # op.add_argument("--headless") # Används ej i nuläget, men den används för att inte visa Edge. och allt görs i bakgrunden.
    op.add_argument("--window-size=%s" % WINDOW_SIZE)
    op.add_experimental_option('excludeSwitches', ['enable-logging']) # Vet ej vad denna gör men den gör så att det fungerar bättre :)

    global driver
    driver = webdriver.Edge(service=ser, options=op)

    global action
    action = ActionChains(driver)


def search_hsa_id(hsa):
    """
    Söker upp hsa-id i EK
    """
    global first_name
    global log_string
    global last_name

    url_ek = "https://www.ek.sll.se/ekadmin/NEIDMgmt"
    # Går in i EK

    try:
        driver.get(url_ek)
    except TimeoutException:
        log_string += "- Något gick fel vid öppnandet av EK. Inget gjordes. \n"
    
    
    try:
        # Försöker att Klickar på sök knappen
        WebDriverWait(driver, 4).until(
            ec.presence_of_element_located((By.LINK_TEXT, "Sök"))
        ).click()
    except:
        # Om sökknappen inte hittas så klickar den på logga in med eTjänstekort.
        WebDriverWait(driver, 5).until(
            ec.presence_of_element_located((By.XPATH, '//span[text()="eTjänstekort"]/..'))
        ).click()

        # Och sedan försöker klicka på sök igen.
        WebDriverWait(driver, 5).until(
            ec.presence_of_element_located((By.LINK_TEXT, "Sök"))
        ).click()

    # Skriver in HSA-IDt som har angets som input till funktionen och klickar enter för att söka
    hsaid_input = WebDriverWait(driver, 7).until(
        ec.presence_of_element_located((By.ID, "hsaIdentity"))
    )
    hsaid_input.send_keys(hsa)
    hsaid_input.send_keys(Keys.RETURN)
    sleep(1)

def change_menyval(menyval):
    # Klicka på Menyval i redigera fönstret
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "dijit_MenuBar_0"))
    ).click()

    # Går till "Person" i menyvalet
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.XPATH, f"//td[text()='{menyval}']/.."))
    ).click()

    sleep(0.2)

def open_personpost(workplace, workplace_input, hsa_id_input, close_redigera_window_after_funk, personpost_position=None):
    """
    Öppnar personposten och tar fram förnam och efternamn på användaren.

    close_redigera_window tar endast emot True or False för att kolla om man vill stänga redigera fönstret."
    """
    global first_name
    global last_name
    global log_string
    global prod
    global chosen_personpost

    chosen_personpost = None
    kk_workplace = ""

    search_hsa_id(hsa_id_input)

    # Tar fram alla sök resultat i en lista
    dojo_grid_row = driver.find_elements(By.CLASS_NAME, "dojoxGridRowTable")

    # Om en personpost position har följt med i callet till funktionen så väljer den den positionen direkt,
    # och hoppar över skanningarna av personposterna.
    if not personpost_position:

        # Om inget HSA-ID hittas i EK. Första positionen är raden för namnen på varje kolumn, ex "status", "e-post", "HSA-ID"
        if len(dojo_grid_row) == 1:
            log_string += f"- Hittade inte {hsa_id_input} i EK \n"
            root.after(10, print_text_in_text_box, f"Hittade inte {hsa_id_input} i EK \n")
            raise NoUserFoundException

        # Går igenom sök resultaten och kollar om arbetsplatsen matchar med det som man har gett som input
        # Går sedan in på den som matchar
        
        personpost_positions = []

        # Hållare för tillgängliga personposters positioner
        available_personposts_not_in_vo = []

        for index, i in enumerate(dojo_grid_row):
            # Hoppar över första loopen för det är raden för namnen på varje kolumn, ex "status", "e-post"
            if index == 0:
                pass
            else:
                hover_cell = i.find_element(By.XPATH, ".//td[2]")
                action.move_to_element(hover_cell).perform()
                sleep(0.4)

                ek_workplace = driver.find_element(
                    By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text  
                
                if workplace in ek_workplace:
                    personpost_positions.append(index)

                    if workplace_input == "kk":
                        kk_workplace_ward = ek_workplace.split(", ")[1]
                        if kk_workplace_ward in kk_gynekologen_enheter:
                            kk_workplace = "gyn"
                            root.after(10, print_text_in_text_box, "Användaren jobbar på KK Gynekologen")
                        elif kk_workplace_ward in kk_obstetriken_enheter:
                            kk_workplace = "obst"
                            root.after(10, print_text_in_text_box, "Användaren jobbar på KK Obstetriken")
                
                elif "Södersjukhuset AB" in ek_workplace:
                    # Här sparas de personposter som är utanför VO för att sedan loopas igenom och kolla om man vill fortsätta med den 
                    available_personposts_not_in_vo.append(index)

        if len(personpost_positions) == 1:
            personpost_position = personpost_positions[0]


        # Kollar vilken personpost man vill fortsätta med. Ifall det är 2 personposter på samma VO eller under någon annan enhet på SÖS.
        elif len(available_personposts_not_in_vo) != 0 or len(personpost_positions) > 1:
            # Om man väljer att fortsätta med personpost utanför VO, sparas positionen på personposten i chosen_personposts_not_in_vo

            for x in available_personposts_not_in_vo:
                continue_with_personpost = messagebox.askyesno(
                    f"Hittade ingen personpost som matchar", f"Hittade ingen personpost som matchar. Vill du fortsätta med personposten på plats {x} i EK? Räknat uppefrån och ner i listan på personposter användaren har."
                )
                if continue_with_personpost:
                    personpost_position = x

                    break
            else:
                for x in personpost_positions:
                    continue_with_personpost = messagebox.askyesno(
                        f"Har flera personposter på samma VO ", f"Användaren har flera personposter på samma VO. Vill du fortsätta med personposten på plats {x} i EK? Räknat uppefrån och ner i listan på personposter användaren har."
                    )

                    if continue_with_personpost:
                        personpost_position = x
                        break
            
            if  not personpost_position:
                log_string += f"- Ingen användare hittades som matchar HSA-ID och arbetsplats. \n"
                root.after(10, print_text_in_text_box, "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
                raise NoUserFoundException

        elif len(personpost_positions) == 0 and len(available_personposts_not_in_vo) == 0:
            log_string += f"- Ingen användare hittades som matchar HSA-ID och arbetsplats. \n"
            root.after(10, print_text_in_text_box, "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
            raise NoUserFoundException

        else:
            root.after(10, print_text_in_text_box, "Något gick fel när personposten skulle öppnas.")
            raise NoUserFoundException
        

    if personpost_position:
        hover_cell = dojo_grid_row[personpost_position].find_element(By.XPATH, ".//td[2]")
        action.move_to_element(hover_cell).perform()
        sleep(0.2)
        ek_workplace = driver.find_element(By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text
        
        action.context_click(
            dojo_grid_row[personpost_position]).perform()
        sleep(0.4)


        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, "//td[text()='Redigera']"
            ))
        ).click()
    
    else:
        root.after(10, print_text_in_text_box, "Något gick fel med öppnandet av personposten!")
        raise NoUserFoundException

    # Byt fokus till nya fönstret
    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    get_name()

    if close_redigera_window_after_funk:
        root.after(0, print_text_in_text_box, "Tagit fram förnamn och efternamn från EK")
        driver.close()
        driver.switch_to.window(window_before)

        return kk_workplace, personpost_position, close_redigera_window_after_funk, window_before, ek_workplace
    
    elif not close_redigera_window_after_funk:

        return kk_workplace, personpost_position, close_redigera_window_after_funk, window_before, ek_workplace
    
def get_name():
    """
    Tar fram förnamn och efternamn o man är på förstasidan inne
    på personposten i EK och sparar ner det i globala variabler.
    """
    global first_name
    global last_name

    # Sparar Namn
    first_name = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.XPATH, '//div[text()="Tilltalsnamn:"]/../following-sibling::input'))
    ).get_attribute("value")

    last_name = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.XPATH, '//div[text()="Efternamn:"]/../following-sibling::input'))
    ).get_attribute("value")
    

def add_pipemail_role(user_titel, window_before):
    """
    Lägger till behörigheten för rörpost inne på personposten.        
    """
    global log_string

    change_menyval("Behörigheter och roller")

    # klicka på Lägg till
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.ID, "NEIDMgmtHTMLParameterNr61addBtn"
        ))
    ).click()

    # klicka på "--Välj Domän--" och sedan väljer rörpost
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.ID, "NEIDMgmtHTMLParameterNr61systemDomainsList"))
    ).click()

    sleep(1)

    select_domain = Select(driver.find_element(
        By.ID, "NEIDMgmtHTMLParameterNr61systemDomainsList"))
    select_domain.select_by_visible_text("Rörpost_SOS")

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.XPATH, '//*[@id="NEIDMgmtHTMLParameterNr61systemRoleTable"]/tbody/tr[2]'))
    )

    # Väljer rörpost
    try:

        if user_titel == "lak" or user_titel == "at_lak" or user_titel == "usk" or user_titel == "paramed" or user_titel == "stud_ssk":
            # om det är läkare, paramed eller usk ska allmänbehörighet tilldelas
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, '//*[@id="Allmän behörighet, SödersjukhusetcbxSR" and @type="radio"]'
                ))
            ).click()

        elif user_titel == "ssk" or "barnmorsk" in user_titel:
            # om det är ssk eller barnmorska så ska läkemedelsbehörighet tilldelas
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, '//*[@id="Läkemedelsbehörighet, SödersjukhusetcbxSR" and @type="radio"]'
                ))
            ).click()

        else:
            print("Något gick snett vid tillägg av rörpost!")
            log_string += "Något gick snett vid tillägg av rörpost, rätt titel inte vald."
            return

        # klicka på OK
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID,
                "NEIDMgmtHTMLParameterNr61closeBtnSR"
            ))
        ).click()

        sleep(0.7)

        # Klicka på spara-stäng i prod
        if prod:
            WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID,
                "saveButton"
            ))
        ).click()

        if not prod:
            root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
            driver.close()

        log_string += "- Lagt till rörpost. \n"

        print("Lyckades med att lägga till rörpost på personposten")
        driver.switch_to.window(window_before)

    except TimeoutException:
        log_string += "- Rörpost behörigheten Hittades inte. Redan tilldelad? \n"
        root.after(10, print_text_in_text_box, "Hittade inte rörpost behörigheten, är den redan tilldelad? Kontrollera manuellt")
        driver.close()
        driver.switch_to.window(window_before)

    return


def add_groupcode(hsa_id_input, user_titel, workplace, workplace_input, personpost_position, window_before):


    global kk_obstetriken_enheter
    global kk_gynekologen_enheter
    global log_string

    if user_titel == "lak" or user_titel == "at_lak":
        root.after(10, print_text_in_text_box, "\nKollar förskrivarkod")

        open_personpost(workplace, workplace_input, hsa_id_input, False, personpost_position=personpost_position)

        # Klicka på Menyval i redigera fönstret
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.ID, "dijit_MenuBar_0"))
        ).click()

        # Klicka på Titel
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.ID, "dijit_MenuItem_1_text"))
        ).click()

        forskrivarkod_box = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((By.ID, "dijit_form_TextBox_0"))
        )

        print(forskrivarkod_box.get_attribute('value'))

        if forskrivarkod_box.get_attribute('value') == "9100009 Vikarierande examinerad läkare" or \
        forskrivarkod_box.get_attribute('value') == "9000001 AT-läkare":
            root.after(10, print_text_in_text_box, "Förskrivarkod redan tillagd. Hoppar över.")

            log_string += f"- Har inte lagt till förskrivarkod då det redan fanns. \n"

            driver.close()

        elif forskrivarkod_box.get_attribute('value') == "Saknar personlig förskrivarkod":
            WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, '//*[@id="widget_hsaGroupPrescriptionCodecomboBox"]/div/div[1]'
                    ))
                ).click()

            if user_titel == "lak":
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, "//li[@role='option' and text()='9100009 Vikarierande examinerad läkare']"
                    ))
                ).click()

            if user_titel == "at_lak":
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, "//li[@role='option' and text()='9000001 AT-läkare']"
                    ))
                ).click()

            # Klicka på spara stäng i prod
            global prod
            if prod:
                driver.find_element(By.ID, "applyButton").click()

            if not prod:
                root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
                driver.close()

            # Loggar
            log_string += f"- Lagt till förskrivarkod. \n"

            root.after(10, print_text_in_text_box, "Lagt till förskrivarkod på personposten")


        elif forskrivarkod_box.get_attribute('value') == "Har personlig förskrivarkod":
            root.after(10, print_text_in_text_box, "Användaren har personlig förskrivarkod.")

            log_string += f"- Har inte lagt till förskrivarkod då det redan fanns. \n"

            driver.close()

        else:
            log_string += f"- Förskrivarkod hoppades över. \n"
            driver.close()

        driver.switch_to.window(window_before)


def add_vmu(hsa, hsa_id_input):
    """
    Lägger till Vårdmedarbetaruppdrag via HSA-ID med Vårdmedarbetaruppdragets HSA-ID som input
    """

    global prod
    global log_string

    search_hsa_id(hsa)

    # Högerklickar på Vårdmedarbetaruppdraget och klickar enter för att komma in på redigera
    dojo_grid_row = driver.find_elements(By.CLASS_NAME, "dojoxGridRowTable")
    vmu_name = dojo_grid_row[1].find_element(By.XPATH, ".//td[2]").text

    root.after(10, print_text_in_text_box, f'Lägger till användaren i Vårdmedarbetaruppdrag "{vmu_name}"')

    action.context_click(dojo_grid_row[1]).perform()
    WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, "//td[text()='Redigera']"
                    ))
                ).click()

    try:
        window_before = driver.window_handles[0]
        window_after = driver.window_handles[1]
        driver.switch_to.window(window_after)

    except (NoSuchWindowException, IndexError):
        # Inget nytt fönster hittades om denna exception triggas
        root.after(10, print_text_in_text_box, f"Kunde inte hitta medarbetaruppdrag {hsa}. Kontrollera manuellt i EK.")
        return

    # Klickar på menyval och sen går in på Vårdmedarbetaruppdragets medlemmar
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.ID, "dijit_PopupMenuBarItem_0_text"))
    ).click()

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "dijit_MenuItem_1_text"))
    ).click()

    driver.set_window_size(1920, 1080)

    sleep(0.1)
    
    # Klickar på "Lägg till"
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "dijit_form_Button_2"))
    ).click()

    # Klicka på "Välj vårdgivare som sök databas" så den söker på hela sös
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.XPATH, "//*[@id='firstRow']/img[2]"))
    ).click()

    # Fyll i HSA-ID och klicka på sök
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "hsaid"))
    ).send_keys(hsa_id_input)

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "dijit_form_Button_4"))
    ).click()

    # Kollar om det dyker upp en alert att personen redan finns med i Vårdmedarbetaruppdraget.
    try:
        WebDriverWait(driver, 3).until(ec.alert_is_present(),
                                       'Person ligger redan med i medarbetaruppdraget. Hoppar över.')

        alert = driver.switch_to.alert
        alert_text = alert.text
        alert.accept()
        driver.close()
        driver.switch_to.window(window_before)
        if alert_text == "Inga sökträffar":
            log_string += f"- Hittade inte Vårdmedarbetaruppdrag {hsa.upper()}. Hoppade över. \n"
            root.after(10, print_text_in_text_box, "Hittade inget användaren vid tillägg av Vårdmedarbetaruppdrag")

        else:
            log_string += f'- Ligger redan med i Vårdmedarbetaruppdrag "{vmu_name}". Hoppade över. \n'
            root.after(10, print_text_in_text_box, f'{hsa_id_input.upper()} ligger redan med i "{vmu_name}". Hoppar över.')
            

        return
    except TimeoutException:
        pass

    # Klickar på lägg till
    driver.find_element(By.ID, "dijit_form_Button_6").click()
    sleep(3)

    # Klicka på spara-stäng ---------- Använd endast i prod
    if prod:
        driver.find_element(By.ID, "saveButton").click()

    if not prod:
        root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
        driver.close()

    # Byter fönster för att få fram namnet på vmut
    driver.switch_to.window(window_before)

    # loggar och skriver ut i konsollen
    log_string += f'- Lagt till Vårdmedarbetaruppdrag "{vmu_name}". \n'
    print(f'Lyckades lägga till {hsa_id_input.upper()} i "{vmu_name}".')


# Easteregg: aHR0cHM6Ly93d3cueW91dHViZS5jb20vd2F0Y2g/dj1kUXc0dzlXZ1hjUQ==
# T20gZHUgbHlja2FzIGxpc3RhIHV0IGRldCBow6RyIGbDpXIgZHUgZ8Okcm5hIGtvbnRha3RhIG1pZyAoTWF0dGlhcykgb2NoIHPDpGdhIGF0dCBkdSBrbsOkY2t0ZSBuw7Z0ZW4hIA==

def add_pascal(user_titel, hsa_id_input, vard_och_behandling_vmu_hsa, workplace_input, kk_workplace):
    """
    Lägger till behörighet för pascal
    """

    global log_string

    if user_titel == "lak" or user_titel == "at_lak" or user_titel == "ssk":
        root.after(10, print_text_in_text_box, "\nLägger till behörighet för Pascal")

    if user_titel == "lak":
        if workplace_input == "kk":
            all_units = {
                "gyn": "83L0",
                "obs": "83KZ",
                "sacskydd": "8442"
            }

            for x in all_units.values():
                add_vmu(x, hsa_id_input)

        else:
            add_vmu(vard_och_behandling_vmu_hsa, hsa_id_input)

            if workplace_input in "ane, int, kar, kir, ort":
                add_vmu("73N6", hsa_id_input)
                
    elif user_titel == "at_lak":
        all_units = {
            "aku": "73N6",
            "ane": "7T5P",
            "bild": "7T5Q",
            "hki": "7TKK",
            "inf": "7TKV",
            "int": "73G7",
            "kar": "73MS",
            "kir": "73MR",
            "ort": "71XF",
            "uro": "7Q8W",
            "sac": "7VR0"
        }

        for x in all_units.values():
            add_vmu(x, hsa_id_input)

    if user_titel == "ssk":
        if workplace_input == "kk" and kk_workplace == "obst":
            vard_och_behandling_vmu_hsa = "83KZ"

        elif workplace_input == "kk" and kk_workplace == "gyn":
            vard_och_behandling_vmu_hsa = "83L0"

        elif workplace_input == "kk" and kk_workplace != "gyn" and kk_workplace != "obst":
            root.after(10, print_text_in_text_box, "Något gick fel vid tillägg av behörighet för pascal." + 
                  "Skriptet kunde inte lista ut om användaren skulle ha behörighet till obst eller gyn.")
            log_string += "Något gick fel vid tillägg av behörioghet för pascal. \
                   Skriptet kunde inte lista ut om användaren skulle ha behörighet till obst eller gyn."
            return

        add_vmu(vard_och_behandling_vmu_hsa, hsa_id_input)


def add_frapp(hsa_id_input, user_titel, workplace_input, ek_workplace):
    """
    Lägger till behörighet till frapp om man uppfyller kraven

    """

    if user_titel == "lak" or user_titel == "ssk":
        root.after(10, print_text_in_text_box, "\nKollar om behörighet i frapp ska läggas till")

    if user_titel == "lak":
        
        if workplace_input == "Akuten":
            # B8Z2 är HSA-IDt till "Ambulans-AnnanVG-Läkare tillgång ambulansjournal-" under akuten
            add_vmu("B8Z2", hsa_id_input)
        elif workplace_input == "Kardiologin":
            # 9XNR är HSA-IDt till "Ambulans-AnnanVG-Läkare tillgång ambulansjournal-" under Kardiologin
            add_vmu("9XNR", hsa_id_input)
        else:
            root.after(10, print_text_in_text_box, "Jobbar inte under Akuten eller Kardiologin, lägger inte till behörighet för frapp.")

    elif user_titel == "ssk":
        
        if "Akutmottagningen, Akutvårdssektionen, Sachsska barn- och ungdomssjukhuset" in ek_workplace:
            # 9XNR är HSA-IDt till "Ambulans-AnnanVG-Sjuksköterska tillgång ambulansjournal-" under Sachsska
            add_vmu("BQWW", hsa_id_input)

        elif "Intensivvårdsenheter MIVA-HIA-IMA" in ek_workplace:
            # 9XNT är HSA-IDt till
            # "Ambulans-AnnanVG-Sjuksköterska HIA tillgång ambulansjournal-" under Kardiologi
            add_vmu("9XNT", hsa_id_input)

        elif "Akutmottagningen, Akut, Södersjukhuset AB" in ek_workplace:
            # 9XNT är HSA-IDt till "Ambulans-AnnanVG-Sjuksköterska tillgång ambulansjournal-" under akuten
            add_vmu("B8XZ", hsa_id_input)

        else:
            root.after(10, print_text_in_text_box, "Ska inte ha behörighet till frapp")
            return


def create_lifecare_user(hsa_id_input):
    """
    Funktion för att skapa en användare i LifeCare om inget konto finns. 
    Måste köras direkt efter sök på HSA-ID i add_LifeCare och köras om inget HSA-ID hittas 
    """

    global log_string
    global first_name
    global last_name
    global prod

    root.after(10, print_text_in_text_box, "Konto i LifeCare saknades. Skapar nytt.")
    log_string += "- Konto i LifeCare saknades, skapar nytt konto. \n"

    # Klicka på skapa användare
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, '//*[text()="Skapa användare"]/../..'
        ))
    ).click()

    # Fyll i alla uppgifter i formuläret
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.ID, 'firstName'
        ))
    ).send_keys(first_name)

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.ID, 'lastName'
        ))
    ).send_keys(last_name)

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.ID, 'hsa'
        ))
    ).send_keys("SE2321000016-" + hsa_id_input)

    # Klicka på spara om i prod:
    if prod:
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, '//*[text()="Spara"]/..'
            ))
        ).click()
        pass
    if not prod:
        root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")

    log_string += "- Konto i LifeCare skapat. \n"
    root.after(10, print_text_in_text_box, "Konto skapat")

def search_for_hsa_id_in_lifecare_after_sign_in(hsa_id_input):
        """
        Namnet på funktionen säger sig självt.
        """
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, '//*[@id="nav"]/li[2]/a[@title="Användare"]'
            ))
        ).click()
        
        # Kollar om användaren finns i systemet redan genom att söka på HSA-IDt.
        # Sök fältet
        search_field = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID, 'searchText'
            ))
        )

        # Skriver in i sökfältet och klickar på att den ska söka på inaktiva
        show_inactive_button = WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID, 'showInactive'
            ))
        )
        show_inactive_button.click()

        search_field.send_keys(hsa_id_input)
        search_field.send_keys(Keys.RETURN)

def add_lifecare(user_titel, hsa_id_input, workplace_input, kk_workplace):
    """
    Funktion för att lägga till behörighet för användare i LifeCare
    """

    global log_string
    global prod

    if workplace_input in "inf, int, kar, kir, onk, ort, uro, kk, gyn, obst" and user_titel == "ssk":

        root.after(10, print_text_in_text_box, "\nLägger till behörighet för LifeCare")

        all_units = {
            "inf": "Infektionskliniken-Venhälsan",
            "int": "Internmedicin",
            "kar": "Kardiologi",
            "kir": "Kirurgi",
            "onk": "Onkologiska kliniken",
            "ort": "Ortopedkliniken",
            "uro": "Urologiska kliniken",
            "kk": "Kvinnosjukvård och förlossning",
            "gyn": "Kvinnosjukvård och förlossning",
            "obst": "Kvinnosjukvård och förlossning",
            }

        # Går in på LifeCare
        driver.get(r"https://lifecare.regionstockholm.se/sp/#/commission")

        # Försöker att klicka på "logga in med e-Tjänstekort" i idp inloggningen,
        # om den failar med att hitta elementet så skippa steget för då är man eventuellt inloggad.
        try:

            WebDriverWait(driver, 5).until(
                ec.presence_of_element_located((
                    By.XPATH, '//*[@id="signin-wrapper"]/div/form/div[1]/a'
                ))
            ).click()

        except TimeoutException:
            pass
        
        try:
            WebDriverWait(driver, 5).until(
                ec.presence_of_element_located((
                    By.XPATH, f'//h3[text()="{all_units[workplace_input]} - Södersjukhuset AB"]/../../button'
                ))
            ).click()
        except TimeoutException:
            #Om denna knappen inte hittas är man inte inloggad
            root.after(10, print_text_in_text_box, "Något gick fel vid inloggning i LifeCare, kontrollera manuellt")
            log_string += "Något gick fel med inloggningen i LifeCare, lades inte till. \n"
            return ()
                   
        # Klicka på användarknappen
        
        search_for_hsa_id_in_lifecare_after_sign_in(hsa_id_input)

        # Kollar om det dyker upp något sök resultat och klickar på den om det hittas annars kör den funktionen för att skapa ett nytt konto.
        try:
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, f'//div[text()="SE2321000016-{hsa_id_input}"]'
                ))
            ).click()

        except TimeoutException:
            create_lifecare_user(hsa_id_input)

        # Klickar på Lägg till medarbetaruppdrag
        try:

            active_status = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, f'//label[text()="Aktiv"]/following-sibling::div'
                ))
            ).text

            if active_status == "Nej":
                raise LifeCareAccountInactive

            add_lifecare_mu(all_units, workplace_input)

        except TimeoutException:
            root.after(10, print_text_in_text_box, "Något gick fel. Kontrollera manuellt")
            log_string += "- Något gick fel i skapandet av LifeCare kontot."

        except LifeCareAccountInactive:
            root.after(10, print_text_in_text_box, "LifeCare kontot är inaktivt. Aktiverar.")
            log_string +="- LifeCare kontot är inaktivt. Aktiverar."

            # Klicka på ändra knappen
            WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, '//span[text()="Ändra"]'
                    ))
                ).click()
            
            # Markera kontot som aktivt
            WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.ID, 'enabled'
                    ))
                ).click()
            
            # Spara om det är prod
            if prod:
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, '//button/span[text()="Spara"]/..'
                    ))
                ).click()

                sleep(1)

                add_lifecare_mu(all_units, workplace_input)

            elif not prod:
                root.after(10, print_text_in_text_box, "Konto inte aktiverat i LifeCare då det inte är prod. Skippar de nästa stegen.")
                log_string += "Konto inte aktiverat i LifeCare då det inte är prod. Skippar de nästa stegen."


def add_lifecare_mu(all_units, workplace_input):
    global log_string
    global prod

    try: 
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, f'//span[text()="Lägg till medarbetaruppdrag"]/../..'
            ))
        ).click()

        # Väljer Proffession och medarbetaruppdrag
        # Väntar på att Proffession och medarbetaruppdrag finns i DOM
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID, "assignment"
            ))
        )

        sleep(0.5)

        select_profession = Select(
            driver.find_element(By.ID, "assignment"))
        select_profession.select_by_visible_text("Sjuksköterska")

        select_mu = Select(driver.find_element(By.ID, "commission"))
        select_mu.select_by_visible_text(
            f"Internt uppdrag - SPU och SIP --- Vårdenhet: {all_units[workplace_input]}")

        # Klickar på spara om i Prod
        if prod:
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, '//*[text()="Spara"]/..'
                ))
            ).click()
            log_string += f"- Lagt till behörighet i lifecare på {all_units[workplace_input]} \n"
        elif not prod:
            root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
            log_string += "- Konto inte skapat för LifeCare då det inte är prod."

        print(
            f"Lagt till behörighet i LifeCare på {all_units[workplace_input]}.")
        root.after(10, print_text_in_text_box, f"Lagt till behörighet i LifeCare på {all_units[workplace_input]}.")
        
    except TimeoutException:
        if not prod:
            root.after(10, print_text_in_text_box, "Konto inte skapat för LifeCare då det inte är prod. Skippar de nästa stegen.")
            log_string += "- Konto inte skapat för LifeCare då det inte är prod. Skippar de nästa stegen."
        else:
            root.after(10, print_text_in_text_box, "Något gick fel. Kontrollera manuellt")
            log_string += "- Något gick fel i skapandet av LifeCare kontot."

def open_personpost_for_removal_of_mu(hsa_id_input, close_redigera_window_after_funk, personpost_position=None):
    """
    Öppnar personposten och tar fram förnam och efternamn på användaren.
    close_redigera_window tar endast emot True or False för att kolla om man vill stänga redigera fönstret.
    personpost_position tar en position i siffer form som den går in per automatik utan att skanna saker.
    """
    global first_name
    global last_name
    global log_string
    global prod
    global chosen_personpost

    chosen_personpost = None

    search_hsa_id(hsa_id_input)
      
    # Tar fram alla sök resultat i en lista
    dojo_grid_row = driver.find_elements(By.CLASS_NAME, "dojoxGridRowTable")

    # Om en personpost position har följt med i callet till funktionen så väljer den den positionen direkt,
    # och hoppar över skanningarna av personposterna.
    personpost_in_sös = False

    for index, i in enumerate(dojo_grid_row):
        # Hoppar över första loopen för det är raden för namnen på varje kolumn, ex "status", "e-post"
        if index == 0:
            pass
        else:
            hover_cell = i.find_element(By.XPATH, ".//td[2]")
            action.move_to_element(hover_cell).perform()
            sleep(0.4)

            ek_workplace = driver.find_element(
                By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]"
            ).text

            if "Södersjukhuset AB" in ek_workplace:
                # Här sparas de om personposter som är utanför SÖS

                root.after(10, print_text_in_text_box, "Hittade en personpost på SÖS.")
                personpost_in_sös = True


    if len(dojo_grid_row) > 1:
        personpost_position = 1
    else:
        log_string += f"- Ingen användare hittades i EK\n"
        root.after(10, print_text_in_text_box, "Ingen användare hittades i EK")
        raise NoUserFoundException

    if personpost_position:
        hover_cell = dojo_grid_row[personpost_position].find_element(By.XPATH, ".//td[2]")
        action.move_to_element(hover_cell).perform()
        sleep(0.2)
        ek_workplace = driver.find_element(By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text
        
        action.context_click(
            dojo_grid_row[personpost_position]
        ).perform()

        sleep(0.4)

        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, "//td[text()='Redigera']"
            ))
        ).click()
    
    else:
        root.after(10, print_text_in_text_box, "Något gick fel med öppnandet av personposten!")
        raise NoUserFoundException

    # Byt fokus till nya fönstret
    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    get_name()

    if close_redigera_window_after_funk or personpost_in_sös:
        root.after(0, print_text_in_text_box, "Tagit fram förnamn och efternamn från EK")
        driver.close()
        driver.switch_to.window(window_before)

        raise NoUserFoundException

    elif not close_redigera_window_after_funk:

        return window_before
    
def get_mu_fom_personpost_and_remove_mu(hsa_id, window_before):
    global log_string
    continue_with_personpost_removal = True

    change_menyval("Vårdmedarbetaruppdrag")

    #Tar fram alla rader i tabellen
    try:

        all_personpost_mu = WebDriverWait(driver, 10).until(
            ec.presence_of_all_elements_located((
                By.XPATH, '//table[@id="admintable"]/tbody/tr/td/a[contains(@onmouseover, "Södersjukhuset AB,Stockholms Läns Landsting")]'
            ))
        )
    except TimeoutException: 
        # Om det inte finns några MUs. fortsätt inte.
        root.after(10, print_text_in_text_box, "Finns inga vårdmedarbetaruppdrag att ta bort.")
        continue_with_personpost_removal = False
        return continue_with_personpost_removal

    # om det är endast ett mu kvar. Fortsätt inte att ta bort MUs då det är inga kvar efter borttag efter den ena. 
    if len(all_personpost_mu) == 1:
        continue_with_personpost_removal= False
    
    # Klickar in på första mu:t och tar bort
    all_personpost_mu[0].click()

    change_menyval("Vårdmedarbetaruppdragets medlemmar")

    # Högerklickar på hsa-id så att man kan söka.
    hsa_id_search_in_mu = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, '//div[contains(@class, "dojoxGridSortNode") and text() = "HSA-id"]'
        ))
    )

    action.context_click(
            hsa_id_search_in_mu
    ).perform()
    
    # Skriver in HSA-ID i sökrutan som dyker upp.
    grid_filter_box = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.ID, 'gridFilterBox'
        ))
    )

    grid_filter_box.send_keys(hsa_id.upper() + Keys.RETURN)
    sleep(1)

    
    # Klickar i bockrutan för att ta bort.
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, '//td/div[contains(@class, "dojoxGridCheckSelector")]'
        ))
    ).click()

    sleep(0.5)

    # Klickar på ta bort
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, '//span[text()="Ta bort"]'
        ))
    ).click()

    sleep(3)

    temp_log_string = f'- Tagit bort behörighet för MU {driver.title.replace("Edit object:", "").replace("- Google Chrome", "")}\n'

    # Om det är prod. Klicka på spara stäng.
    if prod:
        
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.ID,
                "saveButton"
            ))
        ).click()

    if not prod:
        root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")

    log_string += temp_log_string
    root.after(10, print_text_in_text_box, temp_log_string + "\n")
    if not prod:
        driver.close()

    driver.switch_to.window(window_before)


    return continue_with_personpost_removal


def remove_lifecare(user_titel, hsa_id_input):
    """
    Funktion för att lägga till behörighet för användare i LifeCare
    """

    global log_string
    global prod
    global driver

    if user_titel == "ssk":

        root.after(10, print_text_in_text_box, "\nTar bort behörighet för LifeCare")

        # Går in på LifeCare
        driver.get(r"https://lifecare.regionstockholm.se/sp/#/commission")

        # Försöker att klicka på "logga in med e-Tjänstekort" i IDP inloggningen,
        # om den failar med att hitta elementet så skippa steget för då är man eventuellt inloggad.
        try:
            WebDriverWait(driver, 5).until(
                ec.presence_of_element_located((
                    By.XPATH, '//*[@id="signin-wrapper"]/div/form/div[1]/a'
                ))
            ).click()

        except TimeoutException:
            pass
        
        try:
            WebDriverWait(driver, 5).until(
                ec.presence_of_element_located((
                    By.XPATH, f'//h3[text()="Infektionskliniken-Venhälsan - Södersjukhuset AB"]/../../button'
                ))
            ).click()
        except TimeoutException:
            #Om denna knappen inte hittas är man inte inloggad
            root.after(10, print_text_in_text_box, "Något gick fel vid inloggning i LifeCare, kontrollera manuellt")
            log_string += "Något gick fel med inloggningen i LifeCare, lades inte till. \n"
            return ()
                   
        # Klicka på användarknappen
        
        search_for_hsa_id_in_lifecare_after_sign_in(hsa_id_input)

        sleep(10)
       
        try:
            # ----- Kollar om den hittar kontot
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, f'//div[text()="SE2321000016-{hsa_id_input.upper()}"]'
                ))
            ).click()

        except TimeoutException:
            # ----- Om den inte hittar HSA-IDt är kontot redan borttaget!
            root.after(10, print_text_in_text_box, "LifeCare kontot är redan borttaget")
            log_string += "- LifeCare kontot är redan borttaget."

            return
        
        # Tar fram en lista på alla medarbetaruppdrag som finns på användaren.
        try:
            all_lifecare_mu = WebDriverWait(driver, 10).until(
                ec.presence_of_all_elements_located((
                    By.XPATH, "//div[contains(@class, 'accordion-item')]"
                ))
            )
        except TimeoutException:
            # ----- Om den inte hittar något MU är det redan borttaget eller inte haft något
            root.after(10, print_text_in_text_box, "LifeCare kontot är redan borttaget eller har inte haft någon behörighet till sös.")
            log_string += "- LifeCare kontot är redan borttaget."
            return
            
        all_sös_lifecare_mu = []

        for i in all_lifecare_mu:
            lifecare_mu_name = WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, '//h2/button/span/div/label'
                ))
            ).text

            if "- Södersjukhuset AB" in lifecare_mu_name:
                all_sös_lifecare_mu.append(lifecare_mu_name)

        for i in all_sös_lifecare_mu:
            remove_lifecare_sös_mu(i, hsa_id_input)

def remove_lifecare_sös_mu(lifecare_mu, hsa_id_input):
    """
    Funktion för att ta bort MU i Lifecare på en användare. Tar hela namnet på lifecare_mu och hsaidt. tex ("Akut - Södersjukhuset AB", "1234")
    """

    global prod
    global log_string

    if "(INAKTIV)" in lifecare_mu:
        return

    driver.get(r"https://lifecare.regionstockholm.se/sp/#/commission")

    # Försöker att klicka på "logga in med e-Tjänstekort" i IDP inloggningen,
    # om den failar med att hitta elementet så skippa steget för då är man eventuellt inloggad.
    try:
        WebDriverWait(driver, 5).until(
            ec.presence_of_element_located((
                By.XPATH, '//*[@id="signin-wrapper"]/div/form/div[1]/a'
            ))
        ).click()

    except TimeoutException:
        pass

    WebDriverWait(driver, 5).until(
        ec.presence_of_element_located((
            By.XPATH, f'//h3[text()="{lifecare_mu}"]/../../button'
        ))
    ).click()

    search_for_hsa_id_in_lifecare_after_sign_in(hsa_id_input)

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, f'//div[text()="SE2321000016-{hsa_id_input}"]'
        ))
    ).click()

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, f'//span[text()="Ändra"]/../..'
        ))
    ).click()

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, f"//div[contains(@class, 'accordion-item')]/h2/button/span/div/label[text()=' {lifecare_mu} ']" # Måste vara space runt lifecare_mu för att det är så i HTMLen
        ))
    ).click()

    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((
            By.XPATH, f"//div[contains(@class, 'accordion-item')]/h2/button/span/div/label[text()=' Kirurgi - Södersjukhuset AB ']/../../../../..//button[@ngbtooltip = 'Ta bort']"
        ))
    ).click()

    if prod:
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, "//ngb-modal-window//button/span[text()='Ja']"
            ))
        ).click()
        root.after(10, print_text_in_text_box, "LifeCare kontot är nu borttaget")
        log_string += "- LifeCare kontot är borttaget"
    else: 
        root.after(10, print_text_in_text_box, "LifeCare kontot är INTE borttaget. Inte prod.")
    sleep(1)


def send_ids_mail(user_titel_full, hsa_id_input, new_ids_account, remove_ids_account=False):
    """
    Skickar mail för nytt konto till IDS7(Ris och Pacs).\n 
    Funktionen är beroende av att förnamn och efternamn tas när rörpost läggs till i EK.
    """
    global first_name
    global last_name
    global log_string
    global prod

    if new_ids_account or remove_ids_account:
               
        root.after(10, print_text_in_text_box, "\nSkapar ett mail för ids7")

        outlook_app = win32.Dispatch("Outlook.Application", CoInitialize())
        outlook_NS = outlook_app.GetNameSpace("MAPI")

        temp_mail_body = "Hej,\n"
        if new_ids_account:
            temp_mail_body += f"Ny användare, {first_name} {last_name} ({hsa_id_input.lower()}) ({user_titel_full}).\n\n"
        elif remove_ids_account:
            temp_mail_body += f"Användare slutar, {first_name} {last_name} ({hsa_id_input.lower()}).\n\n"
        else:
            root.after(10, print_text_in_text_box, "\nNågot gick fel med skapandet av mailet!")
            return
        temp_mail_body += "Med vänliga hälsningar \n\n" +\
        "SÖS Kontocentral \n" +\
        "Region Stockholm \n" +\
        "Telefon: 08-123 180 73 \n" +\
        "E-post: kontocentralen.sf@regionstockholm.se"

        mailitem = outlook_app.CreateItem(0)
        mailitem.Subject = "Ny användare" if new_ids_account else "Användare slutar"
        mailitem.BodyFormat = 1
        mailitem.To = "bild-itbestallning.sodersjukhuset@regionstockholm.se"
        mailitem.SentOnBehalfOfName = "kontocentralen.sf@regionstockholm.se"
        mailitem.Body = temp_mail_body
        if prod:
            mailitem.Send()
            root.after(0, print_text_in_text_box, "IDS7 mail skickat")
            log_string += "- Mail för IDS7 skickat"
        elif not prod:
            mailitem.Display()
            root.after(0, print_text_in_text_box, "IDS7 mail inte skickat, inte i prod")
            log_string += "- Mail för IDS7 inte skickat, inte prod"


        CoUninitialize()

def create_close_mail(user_titel_full, hsa_id_input, case_number, new_or_remove):
        
        global first_name
        global last_name
        
        if case_number:
            outlook_app = win32.Dispatch("Outlook.Application", CoInitialize())
            outlook_NS = outlook_app.GetNameSpace("MAPI")
            

            mailitem = outlook_app.CreateItem(0)
            mailitem.Subject = f"Order {case_number} klar"
            mailitem.BodyFormat = 1
            mailitem.SentOnBehalfOfName = "kontocentralen.sf@regionstockholm.se"

            if new_or_remove == "new":
                mailitem.Body = "Hej,\n\n" +\
                            f"Standardprofil {user_titel_full if user_titel_full else ''} är klar för {first_name if first_name else ''} {last_name if last_name else ''} ({hsa_id_input.upper()}).\n\n" +\
                            "Med vänliga hälsningar \n\n" +\
                            "SÖS Kontocentral \n" +\
                            "Region Stockholm \n" +\
                            "Telefon: 08-123 180 73 \n" +\
                            "E-post: kontocentralen.sf@regionstockholm.se"
                
            elif new_or_remove == "remove":
                mailitem.Body = "Hej,\n\n" +\
                            f"Inaktivering klar för användare {first_name if first_name else ''} {last_name if last_name else ''} ({hsa_id_input.upper()}).\n\n" +\
                            "Med vänliga hälsningar \n\n" +\
                            "SÖS Kontocentral \n" +\
                            "Region Stockholm \n" +\
                            "Telefon: 08-123 180 73 \n" +\
                            "E-post: kontocentralen.sf@regionstockholm.se"
                
            else:
                root.after(10, print_text_in_text_box, "\nGick inte skapa lösningsmail, något gick fel...")
                return            
            
            root.after(10, print_text_in_text_box, "\nMail för lösningsmailet skapat. Du kan behöva öppna det manuellt nere i aktivitetsfältet.\n"
                    "Klistra in mailadressen som mailen ska skickas till manuellt.")

            mailitem.Display()

            CoUninitialize()

def write_log(log_text):
    """
    Funktion för att skriva i loggarna på G: om det är prod och i samma
    mapp om det inte är prod.
    """

    global log_string

    # Öppnar prod log filerna
    if prod:
        if not os.path.exists(f"G:\\Lit\\Servicedesk\\Verktyg\\SÖS-KC valmeny\\loggar\\{datetime.now().year}\\{datetime.now().month}"):
            os.makedirs(
                f"G:\\Lit\\Servicedesk\\Verktyg\\SÖS-KC valmeny\\loggar\\{datetime.now().year}\\{datetime.now().month}")
        f = open(
            f"G:\\Lit\\Servicedesk\\Verktyg\\SÖS-KC valmeny\\loggar\\{datetime.now().year}\\{datetime.now().month}\\log.txt", "a")

    # Öppnar acc log filerna som läggs i samma mapp som skriptet körs i.
    elif not prod:
        if not os.path.exists(f".\\loggar\\{datetime.now().year}\\{datetime.now().month}"):
            os.makedirs(
                f".\\loggar\\{datetime.now().year}\\{datetime.now().month}")
        f = open(
            f".\\loggar\\{datetime.now().year}\\{datetime.now().month}\\log.txt", "a")

    # Skriver i logfilen och sen stänger filen.
    f.write(log_text)
    f.close()

    # resettar log stringen
    log_string = ""


def add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, user_titel_full, hsa_id_input, new_ids_account, case_number):
    """
    Lägger till alla behörigheter
    """

    global log_string
    global first_name
    global last_name
    global times_run
    global handle_of_the_window_before_minimizing

    if times_run == 0:
        open_browser()

    try:
        try:
            driver.switch_to.window(handle_of_the_window_before_minimizing)
            driver.set_window_rect(0, 0)
        except NameError:
            pass

        if user_titel == "läksek" or user_titel == "stud_paramed":
            # Läkarsekreterare ska endast ha IDS mail och Lösnings beskrivningen.
            
            # Tar endast fram namn på open personpost
            open_personpost(workplace, workplace_input, hsa_id_input, True)

            send_ids_mail(user_titel_full, hsa_id_input, new_ids_account)

            create_close_mail(user_titel_full, hsa_id_input, case_number, "new")

        else:
            # För rörpost
            kk_workplace, personpost_position, close_redigera_window_after_open_personpost, window_before, ek_workplace = open_personpost(workplace, workplace_input, hsa_id_input, False)

            if not close_redigera_window_after_open_personpost:
                add_pipemail_role(user_titel, window_before)

            # Förskrivarkod
            add_groupcode(hsa_id_input, user_titel, workplace, workplace_input, personpost_position, window_before)

            # För Pascal
            add_pascal(user_titel, hsa_id_input,
                    vard_och_behandling_vmu_hsa, workplace_input, kk_workplace)

            # För frapp
            add_frapp(hsa_id_input, user_titel, workplace_input, ek_workplace)

            # För lifecare
            add_lifecare(user_titel, hsa_id_input, workplace_input, kk_workplace)

            # Skickar mail för konto i ids7 / Ris & Pacs
            send_ids_mail(user_titel_full, hsa_id_input, new_ids_account)

            # Skapar upp mall för lösningsmailet
            create_close_mail(user_titel_full, hsa_id_input, case_number, "new")

        write_log(log_string)

        root.after(10, print_text_in_text_box, f"\n{hsa_id_input.upper()} Klar")
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window()

    except NoSuchWindowException:
        root.after(10, print_text_in_text_box, "Edge är nerstängt, öppnar på nytt")
        times_run = 0
        add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number)


def remove_permissions(user_titel, user_titel_full, hsa_id_input, case_number, send_ids_removal_mail):
    global log_string
    global first_name
    global last_name
    global times_run
    global handle_of_the_window_before_minimizing

    if times_run == 0:
        open_browser()

    try:
        try:
            driver.switch_to.window(handle_of_the_window_before_minimizing)
            driver.set_window_rect(0, 0)
        except NameError:
            pass
        
        try:
            continue_with_personpost_removal = True
            while continue_with_personpost_removal:

                window_before = open_personpost_for_removal_of_mu(hsa_id_input, False)

                continue_with_personpost_removal = get_mu_fom_personpost_and_remove_mu(hsa_id_input, window_before)
        except NoUserFoundException:
            # Meddelandet som dyker upp hanteras i open_personpost_for_removal_of_mu
            root.after(10, print_text_in_text_box, f"\n{hsa_id_input.upper()} Hoppar över borttag då den har personpost på sös eller saknar personpost.")
            write_log(log_string)
            handle_of_the_window_before_minimizing = driver.current_window_handle
            driver.minimize_window()
            return

        remove_lifecare(user_titel, hsa_id_input)

        if send_ids_removal_mail:
            send_ids_mail(user_titel_full, hsa_id_input, new_ids_account=False, remove_ids_account=True)

        create_close_mail(user_titel_full, hsa_id_input, case_number, "remove")
        
        write_log(log_string)

        root.after(10, print_text_in_text_box, f"\n{hsa_id_input.upper()} Klar")
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window()

    except NoSuchWindowException:
        root.after(10, print_text_in_text_box, "Edge är nerstängt, öppnar på nytt")
        times_run = 0
        remove_permissions(user_titel, user_titel_full, hsa_id_input, case_number)



def run(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, user_titel_full, hsa_id_input, new_ids_account, case_number):
    global log_string
    global first_name
    global last_name
    global times_run
    global driver
    global handle_of_the_window_before_minimizing

    log_string += ("\n" + "#"*40 + "\n")
    log_string += f'{datetime.now().strftime("%Y-%m-%d, %H:%M")} - {os.getlogin().upper()} Lägger till behörigheter på HSA-ID {hsa_id_input} under {workplace}\n'

    try:
        add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, user_titel_full, hsa_id_input, new_ids_account, case_number)
    except NoUserFoundException:
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window
        write_log(log_string)
    except MoreThanOneAvailablePersonpost:
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window
        write_log(log_string)


    root.after(0, print_text_in_text_box, "-"*90 + "\n")

    first_name = ""
    last_name = ""
    times_run += 1

def run_remove(user_titel, user_titel_full, hsa_id_input, case_number, send_ids_removal_mail):

    global log_string
    global first_name
    global last_name
    global times_run
    global driver
    global handle_of_the_window_before_minimizing

    log_string += ("\n" + "#"*40 + "\n")
    log_string += f'{datetime.now().strftime("%Y-%m-%d, %H:%M")} - {os.getlogin().upper()} tar bort behörigheter på HSA-ID {hsa_id_input}\n'

    try:
        remove_permissions(user_titel, user_titel_full, hsa_id_input, case_number, send_ids_removal_mail)
    except NoUserFoundException:
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window
        write_log(log_string)
    except MoreThanOneAvailablePersonpost:
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window
        write_log(log_string)


    root.after(0, print_text_in_text_box, "-"*90 + "\n")

    first_name = ""
    last_name = ""
    times_run += 1

if __name__ == "__main__":
    # Startar programmet
    print("Skapad av Mattias Pettersson @ Serviceförvaltningen \n")
    get_input()

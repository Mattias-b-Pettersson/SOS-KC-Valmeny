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
from tkinter import messagebox, Scrollbar, VERTICAL
from pythoncom import CoInitialize, CoUninitialize
from sys import exit
import random
import win32com.client as win32
import threading
import customtkinter
import os

#######################################################
# Skapad av Mattias Pettersson @ Serviceförvaltningen #
#######################################################

prod = True
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

    # workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account
    global text_box
    global root

    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("green")

    root = customtkinter.CTk()
    root.title("SÖS-KC Valmeny")
    # root.iconbitmap("icon.ico")
    root.resizable(False, False)

    frame = customtkinter.CTkFrame(master=root)
    frame.grid(row=0, column=0, pady=20, padx=60)

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

    main_label = customtkinter.CTkLabel(master=frame, text="SÖS-KC Valmeny", font=("Roboto", 32))
    main_label.grid(row=0, column=0, columnspan=2, pady=(30), padx=60)

    hsa_id_label = customtkinter.CTkLabel(master=frame, text="Ange ett HSA-ID:", font=("Roboto", 14))
    hsa_id_label.grid(row=1, column=0, pady=0, padx=(30,10))
    hsa_id_entry = customtkinter.CTkEntry(master=frame, placeholder_text="HSA-ID")
    hsa_id_entry.grid(row=2, column=0, pady=(0,20), padx=(30,10))
    
    titel_label = customtkinter.CTkLabel(master=frame, text="Vilken titel har användaren?", font=("Roboto", 14))
    titel_label.grid(row=3, column=0,pady=0, padx=(30,10))
    titel_entry = customtkinter.CTkComboBox(
        master=frame,
        state="readonly",
        values=[
            "Läkare",
            "AT-Läkare",
            "Sjuksköterska",
            "Undersköterska",
            "Paramedicinerare",
            "Barnmorska",
        ],
    )

    titel_entry.grid(row=4, column=0, pady=(0,20), padx=(30,10))

    workplace_label = customtkinter.CTkLabel(master=frame, text="Vart jobbar användaren?", font=("Roboto", 14))
    workplace_label.grid(row=5, column=0, pady=0, padx=(30,10))
    workplace_entry = customtkinter.CTkComboBox(
        master=frame,
        state="readonly",
        values=[
            "Akuten",
            "Anestesin/IVA",
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

    ids_entry = customtkinter.CTkSwitch(master=frame, text="Skicka mail för IDS?", onvalue=True, offvalue=False)
    ids_entry.grid(row=7, column=0, pady=0, padx=(30,10))

    case_number_label = customtkinter.CTkLabel(master=frame, text="Ange ärendenummret:", font=("Roboto", 14))
    case_number_label.grid(row=8, column=0, pady=0, padx=(30,10))
    case_number_entry = customtkinter.CTkEntry(master=frame, placeholder_text="Ärendenummer")
    case_number_entry.grid(row=9, column=0, pady=(0,20), padx=(30,10))
    
    run_button = customtkinter.CTkButton(
        master=frame, 
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
    run_button.grid(row=10, column=0, pady=(10,20), padx=(30,10))

    text_box = customtkinter.CTkTextbox(master=frame, activate_scrollbars=True, width=800, height=400, font=("Roboto", 18))
    text_box.grid(row= 1, rowspan=9, column=1, pady=(3,20), padx=(0,30))
    text_box.configure(state="disabled")

    loading_message = customtkinter.CTkLabel(master=frame, font=("Roboto", 14), text="", width=600)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

    
    

def run_button_func(workplace_input, user_titel, hsa_id_input, new_ids_account, run_button, loading_message, case_number, case_number_entry, hsa_id_entry):

    hsa_id_input = hsa_id_input.upper().replace(" ", "")

    if len(hsa_id_input) != 4:
        messagebox.showerror("SÖS-KC Valmeny error", "HSA-ID inte 4 karaktärer långt. Gör om gör rätt")
        return
    
    if user_titel == "":
        messagebox.showerror("SÖS-KC Valmeny error", "Ingen titel vald. Gör om gör rätt")
        return
    
    if workplace_input == "":
        messagebox.showerror("SÖS-KC Valmeny error", "Inget VO valt. Gör om gör rätt")
        return
    
    if case_number == "":
        make_sure_no_case_number = messagebox.askyesno("Är du säker?", "Är du säker på att du inte vill skapa ett lösningsmail?\n"
                                                       "Ärendenummer inte ifyllt")
        if make_sure_no_case_number:
            root.after(10, print_text_in_text_box, "Skippar att skapa ett lösningsmail")
        else:
            return

    if not new_ids_account:
        if not workplace_input == "Obstetriken" and not user_titel == "Barnmorska":
            make_sure_ids = messagebox.askyesno("Är du säker?", "Är du säker på att du inte vill skicka ett mail för IDS?")
            if not make_sure_ids:
                return
            elif make_sure_ids:
                root.after(10, print_text_in_text_box, "Skippar att skicka IDS mail...")

    if workplace_input == "Obstetriken" and user_titel == "Barnmorska":
        new_ids_account = False

    run_button.configure(state="disabled", bg_color="grey")

    loading_message.configure(text="Jobbar på behörigheterna...")
    loading_message.grid(row= 10, column=1, pady=(0,20), padx=30)

    t = threading.Thread(target=handle_input, args=(workplace_input, user_titel, hsa_id_input, new_ids_account, case_number))
    t.start()
    schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry)

def schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry):
    """
    Schedule the execution of the `check_if_done()` function after
    one second.
    """
    global root
    root.after(500, check_if_done, t, run_button, loading_message, case_number_entry, hsa_id_entry)


def check_if_done(t, run_button, loading_message, case_number_entry, hsa_id_entry):
    global times_checked_if_done
    # If the thread has finished, re-enable the button.

    if not t.is_alive():
        run_button.configure(state="normal", bg_color="green")
        loading_message.grid_forget()
        case_number_entry.delete(0, "end")
        hsa_id_entry.delete(0, "end")

    else:
        # Otherwise check again after one second.
        schedule_check(t, run_button, loading_message, case_number_entry, hsa_id_entry)

def print_text_in_text_box(text):
    global root
    text_box.configure(state="normal")
    text_box.insert(customtkinter.INSERT, f"{text}\n")
    text_box.configure(state="disabled")
    text_box.see(customtkinter.END)

def on_closing():
    global root
    try:   
        driver.close()
    except:
        pass
    exit()
        

def handle_input(workplace_input, user_titel, hsa_id_input, new_ids_account, case_number):

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

    workplace_dictonary = {
        "Akuten": "aku",
        "Anestesin/IVA": "ane",
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
        "ane": "Anestesi IVA, Södersjukhuset AB",
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

    run(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number)


def open_browser():
    # öppnar chrome
    ser = EdgeService(EdgeChromiumDriverManager().install())
    ser.creation_flags = CREATE_NO_WINDOW
    WINDOW_SIZE = "1920,1080"
    op = webdriver.EdgeOptions()
    # op.add_argument("--headless")
    op.add_argument("--window-size=%s" % WINDOW_SIZE)
    op.add_experimental_option('excludeSwitches', ['enable-logging'])
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

    # Klickar på sök knappen
    search_button = WebDriverWait(driver, 20).until(
        ec.presence_of_element_located((By.LINK_TEXT, "Sök"))
    )

    search_button.click()

    # Skriver in HSA-IDt som har angets som input till funktionen och klickar enter för att söka
    hsaid_input = WebDriverWait(driver, 7).until(
        ec.presence_of_element_located((By.ID, "hsaIdentity"))
    )
    hsaid_input.send_keys(hsa)
    hsaid_input.send_keys(Keys.RETURN)
    sleep(1)


def add_pipemail_role(workplace, workplace_input, user_titel, hsa_id_input):
    """
    Lägger till rörpost rollen
    """
    global first_name
    global last_name
    global log_string
    global prod
    global chosen_personposts_not_in_vo

    kk_workplace = ""
    search_hsa_id(hsa_id_input)
    print("\nLägger till rörpost..")
    root.after(10, print_text_in_text_box, "Lägger till rörpost..")

    # Tar fram alla sök resultat i en lista
    dojo_grid_row = driver.find_elements(By.CLASS_NAME, "dojoxGridRowTable")

    # Om inget HSA-ID hittas i EK
    if len(dojo_grid_row) == 1:
        log_string += f"- Hittade inte {hsa_id_input} i EK \n"
        print(f"Hittade inte {hsa_id_input} i EK \n")
        root.after(10, print_text_in_text_box, f"Hittade inte {hsa_id_input} i EK \n")
        raise NoUserFoundException

    # Går igenom sök resultaten och kollar om arbetsplatsen matchar med det som man har gett som input
    # Går sedan in på den som matchar
    
    personpost_position = 0
    available_personposts = 0

    # {"VO, BLA BLA BLA, Södersjukhuset AB": plats där personposten ligger}
    available_personposts_not_in_vo = []

    for index, i in enumerate(dojo_grid_row):
        # Hoppar över första loopen för det är raden för namnen på varje spalt, ex "status", "e-post"
        if index == 0:
            pass
        else:
            hover_cell = i.find_element(By.XPATH, ".//td[2]")
            action.move_to_element(hover_cell).perform()
            sleep(0.4)

            ek_workplace = driver.find_element(
                By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text  
            
            if workplace in ek_workplace:
                personpost_position = index
                available_personposts += 1

                if workplace_input == "kk":
                    kk_workplace_ward = ek_workplace.split(", ")[1]
                    if kk_workplace_ward in kk_gynekologen_enheter:
                        kk_workplace = "gyn"
                        print("Användaren jobbar på KK Gynekologen")
                        root.after(10, print_text_in_text_box, "Användaren jobbar på KK Gynekologen")
                    elif kk_workplace_ward in kk_obstetriken_enheter:
                        kk_workplace = "obst"
                        print("Användaren jobbar på KK Obstetriken")
                        root.after(10, print_text_in_text_box, "Användaren jobbar på KK Obstetriken")
            
            if "Södersjukhuset AB" in ek_workplace:
                available_personposts_not_in_vo.append(index)

    if available_personposts > 1:
        log_string += f"- Finns 2 matchningar i EK för {hsa_id_input} \n"
        print(f"{hsa_id_input} har 2 personposter under valt VO. dubbelkolla manuellt. \n")
        root.after(10, print_text_in_text_box, f"{hsa_id_input} har 2 personposter under valt VO. dubbelkolla manuellt. \n")
        raise MoreThanOneAvailablePersonpost
    
    elif available_personposts == 1:
        action.context_click(
            dojo_grid_row[personpost_position]).perform()
        sleep(0.4)
        WebDriverWait(driver, 10).until(
            ec.presence_of_element_located((
                By.XPATH, "//td[text()='Redigera']"
            ))
        ).click()

    elif len(available_personposts_not_in_vo) != 0:
        chosen_personposts_not_in_vo = 0

        for x in available_personposts_not_in_vo:
            continue_with_personpost_not_in_vo = messagebox.askyesno(
                f"Hittade ingen personpost som matchar", f"Hittade ingen personpost som matchar. Vill du fortsätta med personposten på plats {x} i EK? Räknat uppefrån och ner i listan på personposter användaren har."
            )
            if continue_with_personpost_not_in_vo:
                chosen_personposts_not_in_vo = x

                action.context_click(
                    dojo_grid_row[chosen_personposts_not_in_vo]).perform()
                sleep(0.4)
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, "//td[text()='Redigera']"
                    ))
                ).click()
        
        if chosen_personposts_not_in_vo == 0:
            log_string += f"- Ingen användare hittades som matchar HSA-ID och arbetsplats. \n"
            print(
                "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
            root.after(10, print_text_in_text_box, "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")

            raise NoUserFoundException

    elif available_personposts == 0 and len(available_personposts_not_in_vo) == 0:
        log_string += f"- Ingen användare hittades som matchar HSA-ID och arbetsplats. \n"
        print(
            "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
        root.after(10, print_text_in_text_box, "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")

        raise NoUserFoundException

    else:
        print("Något gick fel när personposten skulle öppnas.")
        root.after(10, print_text_in_text_box, "Något gick fel när personposten skulle öppnas.")
        raise NoUserFoundException

    # Byt fokus till nya fönstret
    window_before = driver.window_handles[0]
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)

    # Sparar Namn till senare ifall om det behövs
    first_name = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "NEIDMgmtHTMLParameterNr24"))
    ).get_attribute("value")

    last_name = WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "NEIDMgmtHTMLParameterNr18"))
    ).get_attribute("value")

    # Klicka på Menyval i redigera fönstret
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located((By.ID, "dijit_MenuBar_0"))
    ).click()
    sleep(0.7)

    # klicka på behörigheter och roller
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.XPATH, "//td[text()='Behörigheter och roller']"))
    ).click()

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
            (By.ID, "NEIDMgmtHTMLParameterNr61systemDomainsList"))
    ).click()
    WebDriverWait(driver, 10).until(
        ec.presence_of_element_located(
            (By.XPATH, '//*[@id="NEIDMgmtHTMLParameterNr61systemRoleTable"]/tbody/tr[2]'))
    )

    # Väljer rörpost
    try:

        if user_titel == "lak" or user_titel == "at_lak" or user_titel == "usk" or user_titel == "paramed":
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
            print("Sparar ej. inte i prod")
            root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
            driver.close()

        log_string += "- Lagt till rörpost. \n"

        print("Lyckades med att lägga till rörpost på personposten")
        driver.switch_to.window(window_before)

    except TimeoutException:
        log_string += "- Rörpost behörigheten Hittades inte. Redan tilldelad? \n"

        print("Hittade inte rörpost behörigheten, är den redan tilldelad? ")
        print("kontrollera manuellt\n")
        root.after(10, print_text_in_text_box, "Hittade inte rörpost behörigheten, är den redan tilldelad? Kontrollera manuellt")
        driver.close()
        driver.switch_to.window(window_before)

    return kk_workplace


def add_groupcode(hsa_id_input, user_titel, workplace):
    global kk_obstetriken_enheter
    global kk_gynekologen_enheter
    global log_string
    global chosen_personposts_not_in_vo

    if user_titel == "lak" or user_titel == "at_lak":
        print("\nKollar förskrivarkod")
        root.after(10, print_text_in_text_box, "\nKollar förskrivarkod")

        search_hsa_id(hsa_id_input)

        dojo_grid_row = driver.find_elements(
            By.CLASS_NAME, "dojoxGridRowTable")
        
        # Om man godkännt att köra på en personpost utanför valt VO så kör den på det. Exempelvis externbemanning.
        if chosen_personposts_not_in_vo != 0:
            action.context_click(
                dojo_grid_row[chosen_personposts_not_in_vo]).perform()
            sleep(0.4)
            WebDriverWait(driver, 10).until(
                ec.presence_of_element_located((
                    By.XPATH, "//td[text()='Redigera']"
                ))
            ).click()

        else: 
            # Går igenom sök resultaten och kollar om arbetsplatsen matchar med det som man har gett som input
            # Går sedan in på den som matchar
            personpost_position = 0
            available_personposts = 0
            
            for index, i in enumerate(dojo_grid_row):
                # Hoppar över första loopen för det är raden för namnen på varje spalt, ex "status", "e-post"
                if index == 0:
                    pass
                else:
                    hover_cell = i.find_element(By.XPATH, ".//td[2]")
                    action.move_to_element(hover_cell).perform()


                    ek_workplace = driver.find_element(
                        By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text  
                    
                    
                    if workplace in ek_workplace:
                        personpost_position = index
                        available_personposts += 1

            if available_personposts > 1:
                handle_of_the_window_before_minimizing = driver.current_window_handle
                driver.minimize_window
                write_log(log_string)
                raise MoreThanOneAvailablePersonpost
            elif available_personposts == 0:
                log_string += f"- Ingen användare hittades som matchar HSA-ID och arbetsplats. \n"
                print(
                    "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
                root.after(10, print_text_in_text_box, "Ingen användare hittades som matchar HSA-ID och arbetsplats. Kontrollera manuellt.")
                raise NoUserFoundException
            elif available_personposts == 1:
                action.context_click(
                    dojo_grid_row[personpost_position]).perform()
                sleep(0.4)
                WebDriverWait(driver, 10).until(
                    ec.presence_of_element_located((
                        By.XPATH, "//td[text()='Redigera']"
                    ))
                ).click()
            
            else:
                print("Något gick fel när personposten skulle öppnas.")
                root.after(10, print_text_in_text_box, "Något gick fel när personposten skulle öppnas.")
                raise NoUserFoundException

        # Byt fokus till nya fönstret
        try:
            window_before = driver.window_handles[0]
            window_after = driver.window_handles[1]
            driver.switch_to.window(window_after)
        except (NoSuchWindowException, IndexError):
            # ifall inget nytt fönster hittades
            print("Något gick fel vid tillägg av gruppförskrivarkod, kontrollera manuellt")
            root.after(10, print_text_in_text_box, "Något gick fel vid tillägg av gruppförskrivarkod, kontrollera manuellt")
            return

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
            print("\nFörskrivarkod redan tillagd. Hoppar över.")
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
                print("Sparar ej. inte i prod")
                root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
                driver.close()

            # Loggar
            log_string += f"- Lagt till förskrivarkod. \n"

            print("Lagt till förskrivarkod på personposten")
            root.after(10, print_text_in_text_box, "Lagt till förskrivarkod på personposten")


        elif forskrivarkod_box.get_attribute('value') == "Har personlig förskrivarkod":
            print("\nAnvändaren har personlig förskrivarkod.")
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

    print(f'Lägger till användaren i Vårdmedarbetaruppdrag "{vmu_name}"')
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
        print(
            f"Kunde inte hitta medarbetaruppdrag {hsa}. Kontrollera manuellt i EK.")
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
            print("Hittade inget HSA-ID")
            root.after(10, print_text_in_text_box, "Hittade inget användaren vid tillägg av Vårdmedarbetaruppdrag")

        else:
            log_string += f'- Ligger redan med i Vårdmedarbetaruppdrag "{vmu_name}". Hoppade över. \n'
            print(
                f'{hsa_id_input.upper()} ligger redan med i "{vmu_name}". Hoppar över. \n')
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
        print("Sparar ej. inte i prod")
        root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
        driver.close()

    # Byter fönster för att få fram namnet på vmut
    driver.switch_to.window(window_before)

    # loggar och skriver ut i konsollen
    log_string += f'- Lagt till Vårdmedarbetaruppdrag "{vmu_name}". \n'
    print(f'Lyckades lägga till {hsa_id_input.upper()} i "{vmu_name}".')


def add_pascal(user_titel, hsa_id_input, vard_och_behandling_vmu_hsa, workplace_input, kk_workplace):
    """
    Lägger till behörighet för pascal
    """

    global log_string

    if user_titel == "lak" or user_titel == "at_lak" or user_titel == "ssk":
        print("\nLägger till behörighet för Pascal")
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
            print("Något gick fel vid tillägg av behörighet för pascal. \n"
                  "Skriptet kunde inte lista ut om användaren skulle ha behörighet till obst eller gyn.")
            root.after(10, print_text_in_text_box, "Något gick fel vid tillägg av behörighet för pascal." + 
                  "Skriptet kunde inte lista ut om användaren skulle ha behörighet till obst eller gyn.")
            log_string += "Något gick fel vid tillägg av behörioghet för pascal. \
                   Skriptet kunde inte lista ut om användaren skulle ha behörighet till obst eller gyn."
            return

        add_vmu(vard_och_behandling_vmu_hsa, hsa_id_input)


def add_frapp(hsa_id_input, user_titel, workplace_input):
    """
    Lägger till behörighet till frapp om man uppfyller kraven
    """

    if user_titel == "lak":
        print("\nKollar om behörighet i frapp ska läggas till")
        root.after(10, print_text_in_text_box, "\nKollar om behörighet i frapp ska läggas till")
        
        if workplace_input == "aku":
            # B8Z2 är HSA-IDt till "Ambulans-AnnanVG-Läkare tillgång ambulansjournal-" under akuten
            add_vmu("B8Z2", hsa_id_input)
        elif workplace_input == "kar":
            # 9XNR är HSA-IDt till "Ambulans-AnnanVG-Läkare tillgång ambulansjournal-" under Kardiologin
            add_vmu("9XNR", hsa_id_input)
        else:
            print(
                "Jobbar inte under Akuten eller Kardiologin, lägger inte till behörighet för frapp.\n")
            root.after(10, print_text_in_text_box, "Jobbar inte under Akuten eller Kardiologin, lägger inte till behörighet för frapp.")

    elif user_titel == "ssk":
        print("\nKollar om behörighet i frapp ska läggas till")
        root.after(10, print_text_in_text_box, "\nKollar om behörighet i frapp ska läggas till")
        search_hsa_id(hsa_id_input)
        dojo_grid_row = driver.find_elements(
            By.CLASS_NAME, "dojoxGridRowTable")

        # Går igenom sök resultaten och kollar om man ska ha frapp
        personpost_counter = 1
        for index, i in enumerate(dojo_grid_row):
            # Hoppar över första loopen för det är raden för namnen på varje spalt, ex "status", "e-post"
            if index == 0:
                pass
            else:
                
                hover_cell = i.find_element(By.XPATH, ".//td[2]")
                action.move_to_element(hover_cell).perform()

                ek_workplace = driver.find_element(
                    By.XPATH, "//*[@id='dijit__MasterTooltip_0']/div[1]").text

                if "Akutmottagningen, Akutvårdssektionen, Sachsska barn- och ungdomssjukhuset" in ek_workplace:
                    # 9XNR är HSA-IDt till "Ambulans-AnnanVG-Sjuksköterska tillgång ambulansjournal-" under Sachsska
                    add_vmu("BQWW", hsa_id_input)
                    break

                elif "Intensivvårdsenheter MIVA-HIA-IMA" in ek_workplace:
                    # 9XNT är HSA-IDt till
                    # "Ambulans-AnnanVG-Sjuksköterska HIA tillgång ambulansjournal-" under Kardiologi
                    add_vmu("9XNT", hsa_id_input)
                    break

                elif "Akutmottagningen, Akut, Södersjukhuset AB" in ek_workplace:
                    # 9XNT är HSA-IDt till "Ambulans-AnnanVG-Sjuksköterska tillgång ambulansjournal-" under akuten
                    add_vmu("B8XZ", hsa_id_input)
                    break

                elif index == len(dojo_grid_row) - 1:
                    print(f"Ska inte ha behörighet till frapp")
                    root.after(10, print_text_in_text_box, "Ska inte ha behörighet till frapp")
                    return

                else:
                    personpost_counter = personpost_counter + 1


def create_lifecare_user(hsa_id_input):
    """
    Funktion för att skapa en användare i LifeCare om inget konto finns. Måste köras direkt efter sök på HSA-ID i add_LifeCare och köras om inget HSA-ID hittas 
    """

    global log_string
    global first_name
    global last_name
    global prod

    print("Konto i LifeCare saknades. Skapar nytt.")
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
        print("Sparar ej. inte i prod")
        root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")

    log_string += "- Konto i LifeCare skapat. \n"
    print("Konto skapat")
    root.after(10, print_text_in_text_box, "Konto skapat")


def add_lifecare(user_titel, hsa_id_input, workplace_input, kk_workplace):
    """
    Funktion för att lägga till behörighet för användare i LifeCare
    """

    global log_string
    global prod

    if workplace_input in "inf, int, kar, kir, onk, ort, uro, kk" and user_titel == "ssk":

        print("\nLägger till behörighet för LifeCare")
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
            print("Något gick fel vid inloggning i LifeCare, kontrollera manuellt")
            root.after(10, print_text_in_text_box, "Något gick fel vid inloggning i LifeCare, kontrollera manuellt")
            log_string += "Något gick fel med inloggningen i LifeCare, lades inte till. \n"
            return ()
                   
        # Klicka på användarknappen
        
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
            print("Något gick fel. Kontrollera manuellt")
            root.after(10, print_text_in_text_box, "Något gick fel. Kontrollera manuellt")
            log_string += "- Något gick fel i skapandet av LifeCare kontot."

        except LifeCareAccountInactive:
            print("LifeCare kontot är inaktivt. Aktiverar.")
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
                print(
                    "Konto inte aktiverat i LifeCare då det inte är prod. Skippar de nästa stegen.")
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
            print("Sparar ej. inte i prod")
            root.after(10, print_text_in_text_box, "Sparar ej. inte i prod")
            log_string += "- Konto inte skapat för LifeCare då det inte är prod."

        
        print(
            f"Lagt till behörighet i LifeCare på {all_units[workplace_input]}.")
        root.after(10, print_text_in_text_box, f"Lagt till behörighet i LifeCare på {all_units[workplace_input]}.")
        
    except TimeoutException:
        if not prod:
            print(
                "Konto inte skapat för LifeCare då det inte är prod. Skippar de nästa stegen.")
            root.after(10, print_text_in_text_box, "Konto inte skapat för LifeCare då det inte är prod. Skippar de nästa stegen.")
            log_string += "- Konto inte skapat för LifeCare då det inte är prod. Skippar de nästa stegen."
        else:
            print("Något gick fel. Kontrollera manuellt")
            root.after(10, print_text_in_text_box, "Något gick fel. Kontrollera manuellt")
            log_string += "- Något gick fel i skapandet av LifeCare kontot."


def send_ids_mail(user_titel, hsa_id_input, new_ids_account):
    """
    Skickar mail för nytt konto till IDS7(Ris och Pacs).\n 
    Funktionen är beroende av att förnamn och efternamn tas när rörpost läggs till i EK.
    """
    global first_name
    global last_name
    global log_string
    global prod

    if new_ids_account:
        if user_titel == "lak":
            titel = "Läkare"
        elif user_titel == "at_lak":
            titel = "AT-Läkare"
        elif user_titel == "ssk":
            titel = "Sjuksköterska"
        elif user_titel == "usk":
            titel = "Undersköterska"
        elif user_titel == "paramed":
            titel = "Paramedicinerare"
        elif user_titel == "barnmorsk":
            titel == "Barnmorska"
        
        
        print("\nSkapar ett mail för ids7")
        root.after(10, print_text_in_text_box, "\nSkapar ett mail för ids7")

        outlook_app = win32.Dispatch("Outlook.Application", CoInitialize())
        outlook_NS = outlook_app.GetNameSpace("MAPI")
        

        mailitem = outlook_app.CreateItem(0)
        mailitem.Subject = "Ny användare"
        mailitem.BodyFormat = 1
        mailitem.To = "bild-itbestallning.sodersjukhuset@regionstockholm.se"
        mailitem.SentOnBehalfOfName = "kontocentralen.sf@regionstockholm.se"
        mailitem.Body = "Hej,\n" +\
                        f"Ny användare, {first_name} {last_name} ({hsa_id_input.lower()}) ({titel}).\n\n" +\
                        "Med vänliga hälsningar \n\n" +\
                        "SÖS Kontocentral \n" +\
                        "Region Stockholm \n" +\
                        "Telefon: 08-123 180 73 \n" +\
                        "E-post: kontocentralen.sf@regionstockholm.se"
        if prod:
            mailitem.Send()
            print("IDS7 mail skickat")
            root.after(10, print_text_in_text_box, "IDS7 mail skickat")
            log_string += "- Mail för IDS7 skickat"
        elif not prod:
            mailitem.Display()
            print("IDS7 mail inte skickat, inte i prod")
            root.after(10, print_text_in_text_box, "IDS7 mail inte skickat, inte i prod")
            log_string += "- Mail för IDS7 inte skickat, inte prod"

        CoUninitialize()

def create_close_mail(user_titel, hsa_id_input, case_number):
        
        global first_name
        global last_name

        if user_titel == "lak":
            titel = "Läkare"
        elif user_titel == "at_lak":
            titel = "AT-Läkare"
        elif user_titel == "ssk":
            titel = "Sjuksköterska"
        elif user_titel == "usk":
            titel = "Undersköterska"
        elif user_titel == "paramed":
            titel = "Paramedicinerare"
        elif user_titel == "barnmorsk":
            titel == "Barnmorska"
        
        if case_number:

            outlook_app = win32.Dispatch("Outlook.Application", CoInitialize())
            outlook_NS = outlook_app.GetNameSpace("MAPI")
            

            mailitem = outlook_app.CreateItem(0)
            mailitem.Subject = f"Order {case_number} klar"
            mailitem.BodyFormat = 1
            mailitem.SentOnBehalfOfName = "kontocentralen.sf@regionstockholm.se"
            mailitem.Body = "Hej,\n\n" +\
                            f"Standardprofil {titel} är klar för {first_name} {last_name} ({hsa_id_input}).\n\n" +\
                            "Med vänliga hälsningar \n\n" +\
                            "SÖS Kontocentral \n" +\
                            "Region Stockholm \n" +\
                            "Telefon: 08-123 180 73 \n" +\
                            "E-post: kontocentralen.sf@regionstockholm.se"
            
            print("\nMail för lösningsmailet skapat. Du kan behöva öppna det manuellt nere i aktivitetsfältet.\n"
                "Klistra in mailadressen som mailen ska skickas till manuellt.")
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


def add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number):
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
        # För rörpost
        kk_workplace = add_pipemail_role(workplace, workplace_input, user_titel, hsa_id_input)

        # Förskrivarkod
        add_groupcode(hsa_id_input, user_titel, workplace)

        # För Pascal
        add_pascal(user_titel, hsa_id_input,
                   vard_och_behandling_vmu_hsa, workplace_input, kk_workplace)

        # För frapp
        add_frapp(hsa_id_input, user_titel, workplace_input)

        # För lifecare
        add_lifecare(user_titel, hsa_id_input, workplace_input, kk_workplace)

        # Skickar mail för konto i ids7 / Ris & Pacs
        send_ids_mail(user_titel, hsa_id_input, new_ids_account)

        # Skapar upp mall för lösningsmailet
        create_close_mail(user_titel, hsa_id_input, case_number)

        write_log(log_string)

        print("\nKlart. Väntar på ny input...")
        root.after(10, print_text_in_text_box, f"\n{hsa_id_input} Klar")
        handle_of_the_window_before_minimizing = driver.current_window_handle
        driver.minimize_window()

    except NoSuchWindowException:
        print("Edge är nerstängt, öppnar på nytt")
        root.after(10, print_text_in_text_box, "Edge är nerstängt, öppnar på nytt")
        times_run = 0
        add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number)

    
    

def run(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number):

    global log_string
    global first_name
    global last_name
    global times_run
    global driver
    global handle_of_the_window_before_minimizing

    log_string += ("\n" + "#"*40 + "\n")
    log_string += f'{datetime.now().strftime("%Y-%m-%d, %H:%M")} - {os.getlogin().upper()} Lägger till behörigheter på HSA-ID {hsa_id_input} under {workplace}\n'

    try:
        add_permissions(workplace, workplace_input, vard_och_behandling_vmu_hsa, user_titel, hsa_id_input, new_ids_account, case_number)
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

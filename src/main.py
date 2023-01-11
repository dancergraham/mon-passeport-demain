import ctypes
import datetime
import time

import requests
import winsound

OFFICE_NAME = "Mairie d'ORVAULT"
API_URL = "https://www.rdv-cni.fr/wp-json/wp/v2/users/22?_=1672254649270"
RDV_URL = "https://www.orvault.fr/vie-pratique/toutes-les-demarches/carte-nationale-didentite-passeport-rendez-vous-en-ligne"
EXCEPTIONS = {"13/02/2023",
              "14/02/2023",
              "15/02/2023",
              "16/02/2023",
              "20/02/2023",
              "21/02/2023",
              "22/02/2023",
              "23/02/2023",
              "24/02/2023",
              "19/04/2023",
              "20/04/2023",
              "21/04/2023",
              "26/04/2023",
              "27/04/2023",
              "02/05/2023",
              }


def blocking_alert(message="message"):
    print(message)
    winsound.PlaySound('sound.wav', winsound.SND_FILENAME)

    WS_EX_TOPMOST = 0x40000
    windowTitle = "Vite mon Passeport!"
    # display a message box; execution will stop here until user acknowledges
    ctypes.windll.user32.MessageBoxExW(None, message, windowTitle, WS_EX_TOPMOST)


def get_data() -> str:
    r = requests.get(API_URL)
    return r.json()


def get_orvault_data(base_data):
    office_data = {"hours": {}}
    assert base_data["data_lieux"]["1"]["nom_lieu"] == OFFICE_NAME
    horaires = base_data["data_lieux"]["1"]["bureaux"]["1"]["horaires"]
    for day, hours in horaires.items():
        reserved_hours = {hour[-5:] for hour in hours if hour.startswith("reserved")}
        available_hours = {hour for hour in hours if not hour.startswith("reserved")}
        available_hours.discard("12:30")
        available_hours.difference_update(reserved_hours)
        office_data["hours"][day] = available_hours
    return office_data


def day_from_date(date) -> str:
    """e.g. day_from_date(2) returns 'mercredi'"""
    return ["lundi",
            "mardi",
            "mercredi",
            "jeudi",
            "vendredi",
            "samedi",
            "dimanche",
            ][date.weekday()]


def vite_mon_passeport():
    base_data = get_data()
    office_data = get_orvault_data(base_data)
    for date, reserved_hours in base_data["data_lieux"]["1"]["bureaux"]["1"]["horaires_reserve"].items():
        if date in EXCEPTIONS:
            continue
        dd, mm, yyyy = map(int, date.split("/"))
        if (mm, dd) >= (5, 19):
            continue
        day = day_from_date(datetime.date(yyyy, mm, dd))
        reserved_hours = set(reserved_hours)
        unreserved_hours = office_data["hours"][day].difference(reserved_hours)
        if unreserved_hours:
            return f"Unreserved hours on {date}\n{RDV_URL}"


def mainloop():
    while True:
        print("checking...", end="")
        message = vite_mon_passeport()
        print("Done")
        if message:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS)
            blocking_alert(message=message)
        time.sleep(180)


if __name__ == "__main__":
    mainloop()

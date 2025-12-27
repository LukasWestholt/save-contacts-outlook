import sys
import win32com.client
import ctypes
from tkinter.simpledialog import askstring


def showUserText(title, text, style):
    ##  Styles:
    ##  0 : OK
    ##  1 : OK | Cancel
    ##  2 : Abort | Retry | Ignore
    ##  3 : Yes | No | Cancel
    ##  4 : Yes | No
    ##  5 : Retry | Cancel
    ##  6 : Cancel | Try Again | Continue

    return ctypes.windll.user32.MessageBoxW(0, text, title, style)


def showUserError(text):
    print(text)
    showUserText("Fehler", text, 0)
    return sys.exit(1)


def build_contact(contact):
    try:
        name = msg.SenderName
        mail = msg.SenderEmailAddress
    except ValueError:
        recipients = msg.Recipients
        if len(recipients) != 1:
            showUserError("Kein eindeutiger Sender oder Empfänger gefunden")
        name = recipients[0].Name
        mail = recipients[0].Address
    if "@" in name and not mail:
        mail = name
        name = askstring("Deine Eingabe wird erfordert", f"Wie heißt die Person von Adresse {mail}? Bitte eingeben als <Nachname>, <Vorname>.")
    else:
        name = askstring("Deine Eingabe wird erfordert", f"Wie heißt die Person von Adresse {mail}? Vielleicht {name}? Bitte eingeben als <Nachname>, <Vorname>.")

    if not name or not mail:
        showUserError("Unvollständiger Kontakt! Abbrechen.")
    
    contact.FullName = name
    contact.Email1Address = mail

def print_address_folder(folder, indent=0):
    print(" " * indent + f"- {folder.Name}")
    for sub in folder.Folders:
        print_address_folder(sub, indent + 2)

def get_target_address_folder(root=True, filter: str = "My Contact Book"):
    # Get default Contacts folder
    contacts_root = namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
    if root:
        return contacts_root

    # OPTIONAL: navigate to a subfolder (contact book)
    for folder in contacts_root.Folders:
        if folder.Name == filter:
            return folder
    print("Contact folders:")
    print_address_folder(contacts_root)

if len(sys.argv) < 2:
    showUserError("No email file provided")

# Force UTF-8 for console, so we can print crazy subject messages
sys.stdout.reconfigure(encoding="utf-8")

msg_path = sys.argv[1]
print(msg_path)

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

msg = namespace.OpenSharedItem(msg_path)

print("Subject:", msg.Subject)
print("From:", msg.SenderName)
print("Email:", msg.SenderEmailAddress)
print("Body preview:", str(msg.Body[:100]).replace("\r\n", "\\n").replace("\n", "\\n"))

target_folder = get_target_address_folder()
if target_folder is None:
    raise Exception("Target contact folder not found")

# Create contact
contact = target_folder.Items.Add("IPM.Contact")
build_contact(contact)
contact.Save()
print("Contact saved successfully")

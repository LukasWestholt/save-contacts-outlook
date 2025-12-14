import sys
import os
import win32com.client

if len(sys.argv) < 2:
    print("No email file provided")
    sys.exit(1)

msg_path = sys.argv[1]

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

msg = namespace.OpenSharedItem(msg_path)

print("Subject:", msg.Subject)
print("From:", msg.SenderName)
print("Email:", msg.SenderEmailAddress)
print("Body preview:", msg.Body[:200])

# Get default Contacts folder
contacts_root = namespace.GetDefaultFolder(10)  # 10 = olFolderContacts

def walk_folders(folder, indent=0):
    print(" " * indent + f"- {folder.Name}")
    for sub in folder.Folders:
        walk_folders(sub, indent + 2)

#print("Contact folders:")
#walk_folders(contacts_root)

# OPTIONAL: navigate to a subfolder (contact book)
#target_folder = None
#for folder in contacts_root.Folders:
#    print(folder.Name)
#    if folder.Name == "My Contact Book":
#        target_folder = folder
#        break

target_folder = contacts_root

if target_folder is None:
    raise Exception("Target contact folder not found")

# Create contact
contact = target_folder.Items.Add("IPM.Contact")

#contact.FirstName = "John"
#contact.LastName = "Doe"
contact.FullName = msg.SenderName
contact.Email1Address = msg.SenderEmailAddress
contact.Save()
print("Contact saved successfully")

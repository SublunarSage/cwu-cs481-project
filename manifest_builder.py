# To use:
#     pip install required packages (lxml and InquirerPy)
#     back up your existing manifest
#     run the script and follow the prompts
#     if you have questions/issues, email me
# 
# If it works but it's ugly, then it still works.
# If it doesn't work... I'll work on it.
# -John
# p.s. remember to validate your manifests using npm run validate


from lxml.builder import ElementMaker
import uuid
from lxml import etree

import os

from InquirerPy import prompt
from InquirerPy import inquirer
from InquirerPy.base.control import Choice
from InquirerPy.separator import Separator

# Define the namespace for the OfficeApp XML
nsmap = {
    None: "http://schemas.microsoft.com/office/appforoffice/1.1",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0",
    "ov": "http://schemas.microsoft.com/office/taskpaneappversionoverrides"
}

E = ElementMaker(nsmap=nsmap)

nsmap_bt = {"bt": "http://schemas.microsoft.com/office/officeappbasictypes/1.0"}
BT = ElementMaker(namespace=nsmap_bt["bt"], nsmap=nsmap_bt)


# Generate a new GUID / RESID
def generate_guid():
    return str(uuid.uuid4()).upper()

def generate_resid():
    generated_guid = generate_guid()
    return generated_guid[:32]

def prettyprint(element, **kwargs):
    xml = etree.tostring(element, pretty_print=True, **kwargs)
    print(xml.decode(), end='')

desired = lambda x: x["desired"] == True
is_label = lambda element: element.tag =="Label"
is_officetab = lambda element: element.tag =="OfficeTab"
is_customtab = lambda element: element.tag =="CustomTab"
is_controlgroup = lambda element: element.tag == "Group"
is_control = lambda element: bool((element.tag == "Menu") or (element.tag == "Button"))
# CustomTabStructure, CustomGroupStructure, etc., remain unchanged
sharedruntimestatus = ""


# OfficeApp children
officeapp_structure = None
officeapp_id = generate_guid()
officeapp_version = "1.0.0.0"
officeapp_providername = "Opertools"
officeapp_defaultlocale = "en-US"
officeapp_displayname = "Opertools Template"
officeapp_description = "A set of tools to help in creation and editing of procedures and policies with standard formatting."
officeapp_iconurl = "https://localhost:3000/assets/icon-32.png"
officeapp_highresolutioniconurl = "https://localhost:3000/assets/icon-64.png"
officeapp_supporturl = "https://www.opertools.com/help"
officeapp_appdomains = ["https://www.opertools.com", "https://www.contoso.com"]
officeapp_hosts = ["Document"]

officeapp_defaultsettings = [
    {"name": "SourceLocation", "desired": True, "value": "https://localhost:3000/taskpane.html"},
    {"name": "RequestedWidth", "desired": False, "value": "800"},
    {"name": "RequestedHeight", "desired": False, "value": "600"}
]

officeapp_permissions = "ReadWriteDocument"

def write_manifest():
    tree = etree.ElementTree(officeapp_structure)
    tree.write('manifest.xml', encoding="UTF-8", xml_declaration=True, pretty_print=True, standalone="yes")

officeapp_structure = E.OfficeApp(
    E.Id(officeapp_id),
    E.Version(officeapp_version),
    E.ProviderName(officeapp_providername),
    E.DefaultLocale(officeapp_defaultlocale),
    E.DisplayName(DefaultValue=officeapp_displayname),
    E.Description(DefaultValue=officeapp_description),
    E.IconUrl(DefaultValue=officeapp_iconurl),
    E.HighResolutionIconUrl(DefaultValue=officeapp_highresolutioniconurl),
    E.SupportUrl(DefaultValue=officeapp_supporturl),
    E.AppDomains(*[E.AppDomain(domain) for domain in officeapp_appdomains]),
    E.Hosts(*[E.Host(Name=host) for host in officeapp_hosts]),
    E.DefaultSettings(*[E(item["name"], DefaultValue=item["value"]) for item in filter(desired, officeapp_defaultsettings)]),
    E.Permissions(officeapp_permissions),
    E.VersionOverrides(
        {"{http://www.w3.org/2001/XMLSchema-instance}type": "VersionOverridesV1_0"},        
        E.Hosts(
            E.Host(

                # E.Runtimes(
                #     E.Runtime(resid="Taskpane.Url", lifetime="long")                          
                # ),
                E.DesktopFormFactor(
                    E.GetStarted(
                        E.Title(resid="GetStarted.Title"),
                        E.Description(resid="GetStarted.Description"),
                        E.LearnMoreUrl(resid="GetStarted.LearnMoreUrl")
                    ),
                    E.FunctionFile(resid="Commands.Url"),
                    E.ExtensionPoint(
                        {"{http://www.w3.org/2001/XMLSchema-instance}type": "PrimaryCommandSurface"},        
                        E.OfficeTab(

                            E.Group(
                                E.Label(resid="CommandsGroup.Label"),
                                E.Icon(
                                    BT.Image(size="16", resid = "Icon.16x16"),
                                    BT.Image(size="32", resid = "Icon.32x32"),
                                    BT.Image(size="80", resid = "Icon.80x80")
                                ),
                                E.Control(
                                    {"{http://www.w3.org/2001/XMLSchema-instance}type": "Button"},
    
                                    E.Label(resid = "TaskpaneButton.Label"),
                                    E.Supertip(
                                        E.Title(resid="TaskpaneButton.Label"),
                                        E.Description(resid="TaskpaneButton.Tooltip")
                                    ),
                                    E.Icon(
                                        BT.Image(size="16", resid = "Icon.16x16"),
                                        BT.Image(size="32", resid = "Icon.32x32"),
                                        BT.Image(size="80", resid = "Icon.80x80")
                                    ),
                                    E.Action(
                                        E.TaskpaneId("ButtonId1"),
                                        E.SourceLocation(resid="Taskpane.Url"),
                                        {"{http://www.w3.org/2001/XMLSchema-instance}type": "ShowTaskpane"},        
                                    ), # Action
                                    id="TaskpaneButton"
                                
                                ), # Control
                                id="CommandsGroup"
                            ), # Group
                            id = "TabHome"
                        ) # OfficeTab
                    ) # ExtensionPoint
                ), # DesktopFormFactor
            **{"{http://www.w3.org/2001/XMLSchema-instance}type": "Document"}          
            ) # Host (document)                           
        ), # Hosts
        E.Resources(
            BT.Images(
                BT.Image(id="Icon.16x16", DefaultValue="https://localhost:3000/assets/icon-16.png"),
                BT.Image(id="Icon.32x32", DefaultValue="https://localhost:3000/assets/icon-32.png"),
                BT.Image(id="Icon.80x80", DefaultValue="https://localhost:3000/assets/icon-80.png"),
                # BT.Image(),
                # BT.Image(),
                # BT.Image(),

            ),
            BT.Urls(
                BT.Url(id="GetStarted.LearnMoreUrl", DefaultValue="https://go.microsoft.com/fwlink/?LinkId=276812"),
                BT.Url(id="Commands.Url", DefaultValue="https://localhost:3000/commands.html"),
                BT.Url(id="Taskpane.Url", DefaultValue="https://localhost:3000/taskpane.html")
            ),
            BT.ShortStrings(
                BT.String(id="GetStarted.Title", DefaultValue="Get started with your sample add-in!"),
                BT.String(id="CommandsGroup.Label", DefaultValue="Commands Group"),
                BT.String(id="TaskpaneButton.Label", DefaultValue="Show Taskpane")
            ),
            BT.LongStrings(
                BT.String(id="GetStarted.Description", DefaultValue="Your sample add-in loaded successfully."),
                BT.String(id="TaskpaneButton.Tooltip", DefaultValue="Click to show a taskpane.")
            )
        ), # Resources
        xmlns= "http://schemas.microsoft.com/office/taskpaneappversionoverrides",
    ), # VersionOverrides
    {"{http://www.w3.org/2001/XMLSchema-instance}type": "TaskPaneApp"},        

) # OfficeApp

def set_shared_runtime():
    global sharedruntimestatus
    if sharedruntimestatus == "SET":
        input("Shared runtime is alreade set")
        return
    sharedruntime_element = E.Runtimes(E.Runtime(resid="Taskpane.Url", lifetime="long"))
    vo_host_element = officeapp_structure.xpath("//VersionOverrides/Hosts/Host")[0]
    vo_host_element.insert(0, sharedruntime_element)
    functionfile_element = officeapp_structure.xpath("//FunctionFile")[0]
    functionfile_element.set("resid", "Taskpane.Url")
    sharedruntimestatus = "SET"
    os.system("cls")
    input("Make sure you modify webpack.config.js or webpack.config.ts appropriately.\nPress Enter to continue")

def MainMenu():
    result = None
    while result !="exit":
        result = DisplayMainMenu()
        if result == "quit":
            return

def DisplayMainMenu():
    global sharedruntimestatus
    result = None
    while result !="exit":
        menu=[
            "Edit basic manifest information",
            f"Set shared runtime {sharedruntimestatus}",
            # "Edit Office Tabs",
            "Edit Custom Tabs",
            "Write manifest.xml",
            "Quit"
        ]
        os.system("cls")
        menuselection = inquirer.rawlist(
            message="Select an action:",
            choices = menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()

        if menuselection == menu[0]: # Edit basic info
            EditBasicInfoMenu()
        elif menuselection == menu[1]: # Set shared runtime
            set_shared_runtime()
        elif menuselection == menu[2]: # Edit Office Tabs
            edit_tabs_menu("office")
        elif menuselection == menu[3]: # Edit Custom Tabs
            edit_tabs_menu("custom")
        elif menuselection == menu[4]: # Build Manifest
            write_manifest()
        elif menuselection == menu[5]: # Exit
            response = input("Are you sure you want to exit? \nEnter 'exit' to exit.")
            return response
        else: # error
            print("an error has occurred")
    
def EditBasicInfoMenu():
    global officeapp_id, officeapp_version, officeapp_providername, officeapp_defaultlocale, officeapp_displayname, officeapp_description, officeapp_iconurl, officeapp_highresolutioniconurl, officeapp_supporturl
    result = None 
    while result != "exit":
        menu =[
                f"ID:................{officeapp_id}",
                f"Version:...........{officeapp_version}",
                f"ProviderName.......{officeapp_providername}",
                f"Default Locale.....{officeapp_defaultlocale}",
                f"Display Name.......{officeapp_displayname}",
                f"Description........{officeapp_description}",
                f"Icon Url...........{officeapp_iconurl}",
                "See More", 
                "Exit to Main Menu"
        ]
        os.system("cls")
        menuselection = inquirer.rawlist(
            message="Select an action:",
            choices=menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()
        print(f"You selected {menuselection}")


        if menuselection == menu[0]: # ID
            print("")
            response = input("The OfficeApp ID is an auto-generated GUID. \nIf you want to modify it, enter 'yes'.")
            if response == "yes":
                officeapp_id = input("Enter a custom GUID:\n")

        elif menuselection == menu[1]: # Version
            print("")
            response = input("If you want to modify the version number, enter 'yes'.\n")
            if response == "yes":
                officeapp_version = input("Enter a custom version number:\n")

        elif menuselection == menu[2]: # ProviderName
            print("")
            response = input("The OfficeApp ID is an auto-generated GUID. \nIf you want to modify it, enter 'yes'.\n")
            if response == "yes":
                officeapp_providername = input("Enter a new provider name:\n")

        elif menuselection == menu[3]: # DefaultLocale
            print("")
            response = input("If you want to modify it, enter 'yes'.\n")
            if response == "yes":
                officeapp_defaultlocale = input("Enter a new default locale:\n")

        elif menuselection == menu[4]: # DisplayName
            print()
            response = input("If you want to modify the display name, enter 'yes'.\n")
            if response == "yes":
                officeapp_displayname = input("Enter a new display name:\n")

        elif menuselection == menu[5]: # Description
            print("")
            response = input("If you want to modify the description, enter 'yes'.\n")
            if response == "yes":
                officeapp_description = input("Enter a description:\n")

        elif menuselection == menu[6]: # IconUrl
            print("")
            response = input("Seriously, just edit the manifest directly for this. \nIf you want to modify it, enter 'yes'.\n")
            if response == "yes":
                officeapp_iconurl = input("Enter a new IconUrl:\n")
        elif menuselection == menu[7]:
            result = EditBasicInfoMenuMore(None)
        elif menuselection == menu[8]: # Exit
            result = "exit"

def EditBasicInfoMenuMore(result):
    global officeapp_id, officeapp_version, officeapp_providername, officeapp_defaultlocale, officeapp_displayname, officeapp_description, officeapp_iconurl, officeapp_highresolutioniconurl, officeapp_supporturl
    result = None 
    while result != "exit":
        menu =[
                f"High Res Icon Url..{officeapp_highresolutioniconurl}",
                f"Support Url........{officeapp_supporturl}", 
                "Go Back",
                "Exit to Main Menu"
            ]
        os.system("cls")
        menuselection = inquirer.rawlist(
            message="Select an action:",
            choices=menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()
        print(f"You selected {menuselection}")


        if menuselection == menu[0]: # HighResolutionIconUrl
            print("")
            response = input("Seriously, just edit the manifest directly for this. \nIf you want to modify it, enter 'yes'.\n")
            if response == "yes":
                officeapp_highresolutioniconurl = input("Enter a new HighResolutionIconUrl:\n")

        elif menuselection == menu[1]: # SupportUrl
            print("")
            response = input("Seriously, just edit the manifest directly for this. \nIf you want to modify it, enter 'yes'.\n")
            if response == "yes":
                officeapp_supporturl = input("Enter a new SupportUrl:\n")
        elif menuselection == menu[2]:
            return None
        elif menuselection == menu[3]: # Exit
            return "exit"

def edit_tabs_menu(tab_type):

    result = None 

    while result != "exit":
        if tab_type == "custom":
            tab_element_list = officeapp_structure.xpath("VersionOverrides/Hosts/Host/DesktopFormFactor/ExtensionPoint/CustomTab")
        elif tab_type == "office":
            tab_element_list = officeapp_structure.xpath("VersionOverrides/Hosts/Host/DesktopFormFactor/ExtensionPoint/OfficeTab")

        generated_menu = []

        for element in tab_element_list:
            generated_menu.append(f"Modify {element.tag} {element.get('id')}")
        generated_menu.append("Create new tab")
        generated_menu.append("Exit to main menu")
        os.system("cls")
        menuselection = inquirer.rawlist(
            message="Editing Custom Tabs:",
            choices=generated_menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()
        index = generated_menu.index(menuselection)
        if index < (len(generated_menu)-2): # Edit tabs
            edit_customtab_menu(tab_element_list[index])
        elif menuselection == generated_menu[-2]: # Create New Tab
            if tab_type == "office":
                input ("I'm still working on that.")
            elif tab_type == "custom":
                create_custom_tab()
        elif menuselection == generated_menu[-1]: # Exit
            result = "exit"

def edit_customtab_menu(tab_element):

    result = None 
    while result != "exit":
        menu = []
        for element in filter(is_controlgroup,tab_element):
            menu.append(f"Edit Group {element.get('id')}")
        menu.append("Change tab name")
        menu.append("Create new control group")
        menu.append("Exit to main menu")

        os.system("cls")

        menuselection = inquirer.rawlist(
            message=f"Editing tab {tab_element.get('id')}",
            choices=menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()
        index = menu.index(menuselection)
        if index < (len(menu)-3): # Edit a group
            edit_controlgroup_menu(tab_element[index])
        elif menuselection == menu[-3]: # Change Tab Name
            edit_element_id(tab_element)
        elif menuselection == menu[-2]: # Create new control group
            create_custom_group(tab_element)
        elif menuselection == menu[-1]: # Exit
            result = "exit"

def create_custom_tab():
    os.system("cls")
    extensionpoint_element = officeapp_structure.xpath("//VersionOverrides/Hosts/Host/DesktopFormFactor/ExtensionPoint")[0]
    tab_id = generate_resid()
    tab_label = input ("What do you want the new tab to be called?\n")
    tab_element = E.CustomTab(
        E.Label(resid=tab_id),
        id = tab_label,
        ) # CustomTab
    extensionpoint_element.append(tab_element)

    shortstrings_element = officeapp_structure.xpath("//VersionOverrides/Resources/*[local-name()='ShortStrings']")[0]
    label_string_element = BT.String(
        id=tab_id, DefaultValue=tab_label)
    shortstrings_element.append(label_string_element)

def edit_element_id(element):
    new_name = input("Enter a new name for the component:\n")
    element.set("id", new_name)

    label_resid = element.xpath("Label")[0].get("resid")
    resource_element = officeapp_structure.xpath(f"//*[@id='{label_resid}']")[0]
    resource_element.set("DefaultValue", new_name)

# TODO TODO TODO
def edit_controlgroup_menu(group_element):

    result = None 
    while result != "exit":
        menu = []
        for element in filter(is_control, group_element):
            menu.append(f"Edit {element.tag} {element.get('id')}")
        menu.append("Change group name")
        menu.append("Create new control")
        menu.append("Exit to main menu")

        os.system("cls")

        menuselection = inquirer.rawlist(
            message=f"Editing group {group_element.get('id')}",
            choices=menu,
            default=1,
            multiselect=False,
            validate=lambda result: len(result) > 1,
        ).execute()
        index = menu.index(menuselection)
        if index < (len(menu)-3): # Edit a group
            print(group_element)
            input()

        elif menuselection == menu[-3]: # Change Tab Name
            edit_element_id(group_element)
        elif menuselection == menu[-2]: # Create new control group
            create_control(group_element)
        elif menuselection == menu[-1]: # Exit
            result = "exit"

def create_custom_group(tab_element):
    group_id = generate_resid()
    group_label = input ("What do you want the new group to be called?")
    icon16_resid = generate_resid()
    icon32_resid = generate_resid()
    icon80_resid = generate_resid()

    group_element = E.Group(
        E.Label(resid=group_id),
        E.Icon(
            BT.Image(size="16", resid = icon16_resid),
            BT.Image(size="32", resid = icon32_resid),
            BT.Image(size="80", resid = icon80_resid)
        ),
        id = group_label,
        ) # Customgroup
    label_element = tab_element.find("Label")
    index = (tab_element.index(label_element))
    tab_element.insert(index,group_element)

    resources_element = officeapp_structure.xpath("//VersionOverrides/Resources")[0]
    resources_element[0].append(BT.Image(id=icon16_resid, DefaultValue = "https://localhost:3000/assets/icon-16.png"))
    resources_element[0].append(BT.Image(id=icon32_resid, DefaultValue = "https://localhost:3000/assets/icon-32.png"))
    resources_element[0].append(BT.Image(id=icon80_resid, DefaultValue = "https://localhost:3000/assets/icon-80.png"))
    resources_element[2].append(BT.String(id=group_id, DefaultValue = group_label))

def edit_group_name(group_element):
    new_name = input("Enter a new name for the tab:\n")
    group_element.set("id", new_name)
    label_resid = group_element[1].get("resid")
    resource_element = officeapp_structure.xpath(f"//*[@id='{label_resid}']")[0]
    resource_element.set("DefaultValue", new_name)

# TODO TODO TODO
def create_control(group_element):
    control_id = generate_resid()
    control_label = input ("What do you want the new control to be called?")

    description = input("Enter a description for the control")
    description_resid = generate_resid()

    title = input("Enter title for control")
    title_resid = generate_resid()

    icon16_resid = generate_resid()
    icon32_resid = generate_resid()
    icon80_resid = generate_resid()

    function_name = input("What JS/TS function will this button call?\n")

    control_element = E.Control(
        E.Label(resid=control_id),
        E.Supertip(
            E.Title(
                resid = title_resid
            ),
            E.Description(
                resid = description_resid
            )
        ),
        E.Icon(
            BT.Image(size="16", resid = icon16_resid),
            BT.Image(size="32", resid = icon32_resid),
            BT.Image(size="80", resid = icon80_resid)
        ),
        E.Action(
            E.FunctionName(
                function_name
            ),
            {"{http://www.w3.org/2001/XMLSchema-instance}type": "ExecuteFunction"},        
        ),
        id = control_label,
        **{"{http://www.w3.org/2001/XMLSchema-instance}type": "Button"}          

        ) # Customcontrol

    group_element.append(control_element)

    resources_element = officeapp_structure.xpath("//VersionOverrides/Resources")[0]
    resources_element[0].append(BT.Image(id=icon16_resid, DefaultValue = "https://localhost:3000/assets/icon-16.png"))
    resources_element[0].append(BT.Image(id=icon32_resid, DefaultValue = "https://localhost:3000/assets/icon-32.png"))
    resources_element[0].append(BT.Image(id=icon80_resid, DefaultValue = "https://localhost:3000/assets/icon-80.png"))

    resources_element[1].append(BT.Url(id="ScriptSource.Url", DefaultValue = "https://localhost:3000/taskpane.html"))

    resources_element[2].append(BT.String(id=control_id, DefaultValue = control_label))
    resources_element[2].append(BT.String(id=title_resid, DefaultValue = title))

    resources_element[3].append(BT.String(id=description_resid, DefaultValue = description))










def main():
    DisplayMainMenu()

    # element = officeapp_structure.xpath("//ExtensionPoint")[0]
    # print(element)
    # create_custom_group(element)

    # create_custom_tab()

    # edit_tabs_menu("custom")

    # list_ribbon_tabs()
    # set_shared_runtime()
    # EditBasicInfoMenu()



    # versionoverrides_element = officeapp_structure.find("VersionOverrides")
    # versionoverrides_element.append(hosts_structure)

    write_manifest()

    prettyprint(officeapp_structure)

main()
